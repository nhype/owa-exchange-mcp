"""Folder tools for the Exchange MCP server.

Ports the get-folders.py logic into an MCP tool using OWAClient.
"""

import json

from mcp.server.fastmcp import Context

from exchange_mcp.server import mcp, AppContext
from exchange_mcp.owa_client import OWAClient


_DISTINGUISHED_NAMES = {
    "msgfolderroot", "inbox", "sentitems", "drafts", "deleteditems",
    "junkemail", "outbox", "calendar", "contacts", "tasks", "notes",
    "journal", "searchfolders",
}

_HEADER_TZ = {
    "__type": "JsonRequestHeaders:#Exchange",
    "RequestServerVersion": "Exchange2013",
    "TimeZoneContext": {
        "__type": "TimeZoneContext:#Exchange",
        "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "Russian Standard Time",
        },
    },
}


def _get_client(ctx: Context) -> OWAClient:
    """Extract the OWAClient from the MCP lifespan context."""
    app_ctx: AppContext = ctx.request_context.lifespan_context
    return app_ctx.client


def _folder_id_dict(folder_id: str) -> dict:
    """Build a typed FolderId or DistinguishedFolderId dict."""
    if folder_id.lower() in _DISTINGUISHED_NAMES:
        return {"__type": "DistinguishedFolderId:#Exchange", "Id": folder_id}
    return {"__type": "FolderId:#Exchange", "Id": folder_id}


@mcp.tool()
def check_session(ctx: Context = None) -> str:
    """Check whether the current OWA session is authenticated.

    Makes a lightweight GetFolder call on the inbox. Returns session
    status, the mailbox display name, and cookie file path.

    Returns:
        JSON object with authenticated (bool), mailbox name, and details.
    """
    client = _get_client(ctx)

    payload = {
        "__type": "GetFolderJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "GetFolderRequest:#Exchange",
            "FolderShape": {
                "__type": "FolderResponseShape:#Exchange",
                "BaseShape": "Default",
            },
            "FolderIds": [
                {"__type": "DistinguishedFolderId:#Exchange", "Id": "inbox"}
            ],
        },
    }

    try:
        data = client.request("GetFolder", payload)
    except Exception as e:
        return json.dumps({
            "authenticated": False,
            "error": str(e),
            "cookie_file": str(client.cookie_file),
        })

    for msg in client.extract_items(data):
        if "Folders" in msg:
            folder = msg["Folders"][0]
            return json.dumps({
                "authenticated": True,
                "mailbox": folder.get("DisplayName", ""),
                "unread": folder.get("UnreadCount", 0),
                "cookie_file": str(client.cookie_file),
            })

    return json.dumps({
        "authenticated": False,
        "error": "Unexpected response",
        "cookie_file": str(client.cookie_file),
    })


@mcp.tool()
def get_folders(
    parent_folder_id: str = "msgfolderroot",
    recursive: bool = False,
    ctx: Context = None,
) -> str:
    """List mail folders from the Exchange mailbox.

    Args:
        parent_folder_id: Parent folder to list children of.
            Defaults to "msgfolderroot" (top-level). Can be a
            distinguished folder name or a raw folder ID.
        recursive: If True, traverse all subfolders recursively (Deep).
            If False, only list immediate children (Shallow).

    Returns:
        JSON array of folder objects with: name, id, total_count,
        unread_count, child_folder_count.
    """
    client = _get_client(ctx)

    traversal = "Deep" if recursive else "Shallow"

    # Determine parent folder id type
    # Distinguished folder names are short lowercase strings
    parent_folder = _folder_id_dict(parent_folder_id)

    payload = {
        "__type": "FindFolderJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "FindFolderRequest:#Exchange",
            "FolderShape": {
                "__type": "FolderResponseShape:#Exchange",
                "BaseShape": "Default",
            },
            "ParentFolderIds": [parent_folder],
            "Traversal": traversal,
            "Paging": {
                "__type": "IndexedPageView:#Exchange",
                "BasePoint": "Beginning",
                "Offset": 0,
                "MaxEntriesReturned": 200,
            },
        },
    }

    try:
        data = client.request("FindFolder", payload)
    except Exception as e:
        return json.dumps({"error": str(e)})

    folders = []
    for msg in client.extract_items(data):
        if "RootFolder" in msg and "Folders" in msg["RootFolder"]:
            for f in msg["RootFolder"]["Folders"]:
                folders.append({
                    "name": f.get("DisplayName", "Unknown"),
                    "id": f.get("FolderId", {}).get("Id", ""),
                    "total_count": f.get("TotalCount", 0),
                    "unread_count": f.get("UnreadCount", 0),
                    "child_folder_count": f.get("ChildFolderCount", 0),
                })

    return json.dumps(folders, ensure_ascii=False)


@mcp.tool()
def create_folder(
    name: str,
    parent_folder_id: str = "msgfolderroot",
    ctx: Context = None,
) -> str:
    """Create a new mail folder.

    Args:
        name: Display name for the new folder.
        parent_folder_id: Parent folder to create under.
            Defaults to "msgfolderroot" (top-level). Can be a
            distinguished folder name or a raw folder ID.

    Returns:
        JSON object with the created folder's id and name.
    """
    client = _get_client(ctx)

    payload = {
        "__type": "CreateFolderJsonRequest:#Exchange",
        "Header": _HEADER_TZ,
        "Body": {
            "__type": "CreateFolderRequest:#Exchange",
            "ParentFolderId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": _folder_id_dict(parent_folder_id),
            },
            "Folders": [
                {
                    "__type": "Folder:#Exchange",
                    "DisplayName": name,
                    "FolderClass": "IPF.Note",
                }
            ],
        },
    }

    try:
        data = client.request_header_payload("CreateFolder", payload)
    except Exception as e:
        return json.dumps({"error": str(e)})

    for msg in client.extract_items(data):
        if "Folders" in msg:
            folder = msg["Folders"][0]
            return json.dumps({
                "success": True,
                "name": name,
                "id": folder.get("FolderId", {}).get("Id", ""),
            })

    return json.dumps({"error": "Unexpected response", "raw": str(data)})


@mcp.tool()
def rename_folder(
    folder_id: str,
    new_name: str,
    ctx: Context = None,
) -> str:
    """Rename an existing mail folder.

    Args:
        folder_id: The Exchange folder ID to rename (from get_folders).
        new_name: New display name for the folder.

    Returns:
        JSON object with success status and new folder id.
    """
    client = _get_client(ctx)

    payload = {
        "__type": "UpdateFolderJsonRequest:#Exchange",
        "Header": _HEADER_TZ,
        "Body": {
            "__type": "UpdateFolderRequest:#Exchange",
            "FolderChanges": [
                {
                    "__type": "FolderChange:#Exchange",
                    "FolderId": {
                        "__type": "FolderId:#Exchange",
                        "Id": folder_id,
                    },
                    "Updates": [
                        {
                            "__type": "SetFolderField:#Exchange",
                            "Path": {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "FolderDisplayName",
                            },
                            "Folder": {
                                "__type": "Folder:#Exchange",
                                "DisplayName": new_name,
                            },
                        }
                    ],
                }
            ],
        },
    }

    try:
        data = client.request_header_payload("UpdateFolder", payload)
    except Exception as e:
        return json.dumps({"error": str(e)})

    for msg in client.extract_items(data):
        if "Folders" in msg:
            folder = msg["Folders"][0]
            return json.dumps({
                "success": True,
                "new_name": new_name,
                "id": folder.get("FolderId", {}).get("Id", ""),
            })

    return json.dumps({"error": "Unexpected response", "raw": str(data)})


@mcp.tool()
def empty_folder(
    folder_id: str,
    delete_sub_folders: bool = False,
    permanent: bool = False,
    ctx: Context = None,
) -> str:
    """Empty all items from a mail folder.

    Args:
        folder_id: The Exchange folder ID to empty (from get_folders).
        delete_sub_folders: If True, also delete sub-folders. Default False.
        permanent: If True, permanently delete items (HardDelete).
            Otherwise move to Deleted Items. Default False.

    Returns:
        JSON object with success status.
    """
    client = _get_client(ctx)

    delete_type = "HardDelete" if permanent else "MoveToDeletedItems"

    payload = {
        "__type": "EmptyFolderJsonRequest:#Exchange",
        "Header": _HEADER_TZ,
        "Body": {
            "__type": "EmptyFolderRequest:#Exchange",
            "FolderIds": [
                {"__type": "FolderId:#Exchange", "Id": folder_id}
            ],
            "DeleteType": delete_type,
            "DeleteSubFolders": delete_sub_folders,
            "SuppressReadReceipt": True,
        },
    }

    try:
        data = client.request_header_payload("EmptyFolder", payload)
    except Exception as e:
        return json.dumps({"error": str(e)})

    for msg in client.extract_items(data):
        if msg.get("ResponseClass") == "Success":
            return json.dumps({"success": True, "folder_id": folder_id})

    return json.dumps({"error": "Unexpected response", "raw": str(data)})


@mcp.tool()
def delete_folder(
    folder_id: str,
    permanent: bool = False,
    ctx: Context = None,
) -> str:
    """Delete a mail folder.

    Args:
        folder_id: The Exchange folder ID to delete (from get_folders).
        permanent: If True, permanently delete (HardDelete).
            Otherwise move to Deleted Items. Default False.

    Returns:
        JSON object with success status.
    """
    client = _get_client(ctx)

    delete_type = "HardDelete" if permanent else "MoveToDeletedItems"

    payload = {
        "__type": "DeleteFolderJsonRequest:#Exchange",
        "Header": _HEADER_TZ,
        "Body": {
            "__type": "DeleteFolderRequest:#Exchange",
            "FolderIds": [
                {"__type": "FolderId:#Exchange", "Id": folder_id}
            ],
            "DeleteType": delete_type,
        },
    }

    try:
        data = client.request_header_payload("DeleteFolder", payload)
    except Exception as e:
        return json.dumps({"error": str(e)})

    for msg in client.extract_items(data):
        if msg.get("ResponseClass") == "Success":
            return json.dumps({"success": True, "folder_id": folder_id})

    return json.dumps({"error": "Unexpected response", "raw": str(data)})


@mcp.tool()
def move_folder(
    folder_id: str,
    target_parent_folder_id: str = "msgfolderroot",
    ctx: Context = None,
) -> str:
    """Move a mail folder to a different parent folder.

    Args:
        folder_id: The Exchange folder ID to move (from get_folders).
        target_parent_folder_id: Destination parent folder.
            Defaults to "msgfolderroot" (top-level). Can be a
            distinguished folder name or a raw folder ID.

    Returns:
        JSON object with success status and new folder ID.
    """
    client = _get_client(ctx)

    payload = {
        "__type": "MoveFolderJsonRequest:#Exchange",
        "Header": _HEADER_TZ,
        "Body": {
            "__type": "MoveFolderRequest:#Exchange",
            "FolderIds": [
                {"__type": "FolderId:#Exchange", "Id": folder_id}
            ],
            "ToFolderId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": _folder_id_dict(target_parent_folder_id),
            },
        },
    }

    try:
        data = client.request_header_payload("MoveFolder", payload)
    except Exception as e:
        return json.dumps({"error": str(e)})

    for msg in client.extract_items(data):
        if "Folders" in msg:
            folder = msg["Folders"][0]
            return json.dumps({
                "success": True,
                "folder_id": folder.get("FolderId", {}).get("Id", ""),
            })

    return json.dumps({"error": "Unexpected response", "raw": str(data)})
