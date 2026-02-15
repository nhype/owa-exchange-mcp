"""Email tools for the Exchange MCP server.

Provides tools for reading, sending, replying, forwarding, and managing
emails via the OWA Exchange API.
"""

import json

from mcp.server.fastmcp import Context

from exchange_mcp.server import mcp, AppContext
from exchange_mcp.owa_client import OWAClient, SessionExpiredError
from exchange_mcp.utils import html_to_text, extract_links_from_html


def _get_client(ctx: Context) -> OWAClient:
    """Extract the OWAClient from the MCP lifespan context."""
    app_ctx: AppContext = ctx.request_context.lifespan_context
    return app_ctx.client


def _get_change_key(client: OWAClient, item_id: str) -> str | None:
    """Fetch the ChangeKey for an item via GetItem (IdOnly).

    OWA requires the ChangeKey on write operations like reply/forward.
    """
    payload = {
        "__type": "GetItemJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "GetItemRequest:#Exchange",
            "ItemShape": {
                "__type": "ItemResponseShape:#Exchange",
                "BaseShape": "IdOnly",
            },
            "ItemIds": [{"__type": "ItemId:#Exchange", "Id": item_id}],
        },
    }
    data = client.request("GetItem", payload)
    for msg in client.extract_items(data):
        if "Items" in msg:
            for item in msg["Items"]:
                return item.get("ItemId", {}).get("ChangeKey")
    return None


def _extract_email_summary(item: dict) -> dict:
    """Extract a summary dict from a FindItem result item."""
    item_type = item.get("__type", "Message:#Exchange")
    is_meeting = any(
        t in item_type
        for t in ("MeetingRequest", "MeetingResponse", "MeetingCancellation")
    )

    # Resolve sender from From -> Organizer -> Sender
    from_data = item.get("From", {}).get("Mailbox", {})
    if not from_data:
        from_data = item.get("Organizer", {}).get("Mailbox", {})
    if not from_data:
        from_data = item.get("Sender", {}).get("Mailbox", {})

    email = {
        "subject": item.get("Subject", "(No subject)"),
        "from": from_data.get("EmailAddress", ""),
        "from_name": from_data.get("Name", ""),
        "date": item.get(
            "DateTimeSent",
            item.get("DateTimeReceived", item.get("DateTimeCreated", "")),
        ),
        "is_read": item.get("IsRead", False),
        "has_attachments": item.get("HasAttachments", False),
        "item_id": item.get("ItemId", {}).get("Id", ""),
        "size": item.get("Size", 0),
        "is_meeting": is_meeting,
        "type": "Meeting" if is_meeting else "Email",
        "preview": item.get("Preview", ""),
        "has_links": False,
    }

    # Basic recipients from DisplayTo/DisplayCc
    email["to"] = (
        [t.strip() for t in item.get("DisplayTo", "").split(";") if t.strip()]
        if item.get("DisplayTo")
        else []
    )
    email["cc"] = (
        [c.strip() for c in item.get("DisplayCc", "").split(";") if c.strip()]
        if item.get("DisplayCc")
        else []
    )

    if is_meeting:
        email["location"] = item.get("Location", "")
        email["start"] = item.get(
            "Start", item.get("StartWallClock", item.get("ReminderDueBy", ""))
        )
        email["end"] = item.get("End", item.get("EndWallClock", ""))

    return email


def _get_item_details(client: OWAClient, item_id: str) -> dict:
    """Get full email details (body, recipients, attachments) via GetItem."""
    payload = {
        "__type": "GetItemJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "V2017_08_18",
        },
        "Body": {
            "__type": "GetItemRequest:#Exchange",
            "ItemShape": {
                "__type": "ItemResponseShape:#Exchange",
                "BaseShape": "AllProperties",
                "BodyType": "HTML",
            },
            "ItemIds": [{"__type": "ItemId:#Exchange", "Id": item_id}],
        },
    }

    data = client.request("GetItem", payload)
    result = {
        "item_id": item_id,
        "subject": "",
        "from": "",
        "from_name": "",
        "date": "",
        "to": [],
        "cc": [],
        "bcc": [],
        "body": "",
        "body_type": "Text",
        "is_read": False,
        "has_attachments": False,
        "has_links": False,
        "importance": "Normal",
        "attachments": [],
    }

    for msg in client.extract_items(data):
        if "Items" not in msg:
            continue
        for item in msg["Items"]:
            result["subject"] = item.get("Subject", "(No subject)")

            # Sender
            from_data = item.get("From", {}).get("Mailbox", {})
            if not from_data:
                from_data = item.get("Sender", {}).get("Mailbox", {})
            result["from"] = from_data.get("EmailAddress", "")
            result["from_name"] = from_data.get("Name", "")

            result["date"] = item.get(
                "DateTimeSent",
                item.get("DateTimeReceived", item.get("DateTimeCreated", "")),
            )
            result["is_read"] = item.get("IsRead", False)
            result["has_attachments"] = item.get("HasAttachments", False)
            result["importance"] = item.get("Importance", "Normal")

            # Body
            body_val = item.get("Body", {}).get("Value", "")
            body_type = item.get("Body", {}).get("BodyType", "Text")
            if body_type == "HTML":
                result["has_links"] = bool(extract_links_from_html(body_val))
                result["body"] = html_to_text(body_val)
            else:
                result["body"] = body_val
            result["body_type"] = body_type

            # To recipients
            for r in item.get("ToRecipients", []):
                name = r.get("Name", "")
                addr = r.get("EmailAddress", "")
                if name and addr:
                    result["to"].append(f"{name} <{addr}>")
                elif addr:
                    result["to"].append(addr)

            # CC recipients
            for r in item.get("CcRecipients", []):
                name = r.get("Name", "")
                addr = r.get("EmailAddress", "")
                if name and addr:
                    result["cc"].append(f"{name} <{addr}>")
                elif addr:
                    result["cc"].append(addr)

            # BCC recipients
            for r in item.get("BccRecipients", []):
                name = r.get("Name", "")
                addr = r.get("EmailAddress", "")
                if name and addr:
                    result["bcc"].append(f"{name} <{addr}>")
                elif addr:
                    result["bcc"].append(addr)

            # Attachments (with IDs for download)
            for att in item.get("Attachments", []):
                result["attachments"].append(
                    {
                        "name": att.get("Name", ""),
                        "size": att.get("Size", 0),
                        "content_type": att.get("ContentType", ""),
                        "attachment_id": att.get("AttachmentId", {}).get("Id", ""),
                        "is_inline": att.get("IsInline", False),
                    }
                )

            # Meeting-specific fields
            item_type = item.get("__type", "")
            if any(
                t in item_type
                for t in ("MeetingRequest", "MeetingResponse", "MeetingCancellation", "CalendarItem")
            ):
                result["location"] = item.get(
                    "Location",
                    item.get("EnhancedLocation", {}).get("DisplayName", ""),
                )
                result["start"] = item.get("Start", "")
                result["end"] = item.get("End", "")
                result["required_attendees"] = []
                result["optional_attendees"] = []
                for a in item.get("RequiredAttendees", []):
                    mb = a.get("Mailbox", {})
                    name = mb.get("Name", "")
                    addr = mb.get("EmailAddress", "")
                    if name and addr:
                        result["required_attendees"].append(f"{name} <{addr}>")
                    elif addr:
                        result["required_attendees"].append(addr)
                for a in item.get("OptionalAttendees", []):
                    mb = a.get("Mailbox", {})
                    name = mb.get("Name", "")
                    addr = mb.get("EmailAddress", "")
                    if name and addr:
                        result["optional_attendees"].append(f"{name} <{addr}>")
                    elif addr:
                        result["optional_attendees"].append(addr)

            return result

    return result


def _build_recipient_list(emails: str) -> list[dict]:
    """Build a list of Mailbox dicts from a comma-separated email string.

    NOTE: OWA rejects __type annotations on recipient Mailbox dicts for
    CreateItem (Message). Use plain dicts without __type.
    """
    recipients = []
    for addr in emails.split(","):
        addr = addr.strip()
        if addr:
            recipients.append(
                {
                    "Name": addr,
                    "EmailAddress": addr,
                    "RoutingType": "SMTP",
                }
            )
    return recipients


# ------------------------------------------------------------------
# Tools
# ------------------------------------------------------------------


@mcp.tool()
def get_emails(
    folder: str = "Inbox",
    limit: int = 10,
    offset: int = 0,
    include_body: bool = False,
    unread_only: bool = False,
    ids_only: bool = False,
    ctx: Context = None,
) -> str:
    """Get emails from a mailbox folder.

    Args:
        folder: Folder name (Inbox, Sent, Drafts, Deleted, Junk, or custom name).
        limit: Maximum number of emails to return (default 10, max 50).
        offset: Number of emails to skip for pagination.
        include_body: If True, fetch full body for each email (slower).
        unread_only: If True, only return unread emails.
        ids_only: If True, return only item IDs and dates (compact, for bulk ops).
            Max limit raised to 500 in this mode.
    """
    try:
        client = _get_client(ctx)

        # Clamp limit (higher cap for ids_only)
        max_limit = 500 if ids_only else 50
        if limit > max_limit:
            limit = max_limit

        # Resolve folder name to ID
        folder_id = client.get_folder_id(folder)
        if not folder_id:
            return json.dumps({"error": f"Folder '{folder}' not found."})

        # Build FindItem payload
        if ids_only:
            item_shape = {
                "__type": "ItemResponseShape:#Exchange",
                "BaseShape": "IdOnly",
                "AdditionalProperties": [
                    {
                        "__type": "PropertyUri:#Exchange",
                        "FieldURI": "DateTimeReceived",
                    },
                    {
                        "__type": "PropertyUri:#Exchange",
                        "FieldURI": "Subject",
                    },
                ],
            }
        else:
            item_shape = {
                "__type": "ItemResponseShape:#Exchange",
                "BaseShape": "AllProperties",
            }

        find_body = {
            "__type": "FindItemRequest:#Exchange",
            "ItemShape": item_shape,
            "ParentFolderIds": [
                {"__type": "FolderId:#Exchange", "Id": folder_id}
            ],
            "Traversal": "Shallow",
            "Paging": {
                "__type": "IndexedPageView:#Exchange",
                "BasePoint": "Beginning",
                "Offset": offset,
                "MaxEntriesReturned": limit,
            },
            "SortOrder": [
                {
                    "__type": "SortResults:#Exchange",
                    "Order": "Descending",
                    "Path": {
                        "__type": "PropertyUri:#Exchange",
                        "FieldURI": "DateTimeReceived",
                    },
                }
            ],
        }

        # Add filter for unread only
        if unread_only:
            find_body["Restriction"] = {
                "__type": "RestrictionType:#Exchange",
                "Item": {
                    "__type": "IsEqualTo:#Exchange",
                    "FieldURIOrConstant": {
                        "__type": "FieldURIOrConstantType:#Exchange",
                        "Item": {
                            "__type": "ConstantValueType:#Exchange",
                            "Value": "false",
                        },
                    },
                    "Path": {
                        "__type": "PropertyUri:#Exchange",
                        "FieldURI": "IsRead",
                    },
                },
            }

        payload = {
            "__type": "FindItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "Exchange2013",
            },
            "Body": find_body,
        }

        data = client.request("FindItem", payload)

        items = []
        for msg in client.extract_items(data):
            if "RootFolder" in msg and "Items" in msg["RootFolder"]:
                items = msg["RootFolder"]["Items"]
                break

        if not items:
            return json.dumps(
                {"item_ids": [], "count": 0} if ids_only
                else {"emails": [], "count": 0}
            )

        if ids_only:
            result = []
            for item in items:
                result.append({
                    "item_id": item.get("ItemId", {}).get("Id", ""),
                    "date": item.get("DateTimeReceived", ""),
                    "subject": item.get("Subject", ""),
                })
            return json.dumps({"item_ids": result, "count": len(result)})

        emails = []
        for item in items:
            email = _extract_email_summary(item)

            if include_body and email["item_id"]:
                details = _get_item_details(client, email["item_id"])
                email["to"] = details["to"]
                email["cc"] = details["cc"]
                email["body"] = details["body"]
                email["has_links"] = details.get("has_links", False)

                if email.get("is_meeting"):
                    email["location"] = details.get("location", "")
                    email["start"] = details.get("start", "") or email.get("start", "")
                    email["end"] = details.get("end", "") or email.get("end", "")
                    email["required_attendees"] = details.get("required_attendees", [])
                    email["optional_attendees"] = details.get("optional_attendees", [])

            emails.append(email)

        return json.dumps({"emails": emails, "count": len(emails)})

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to get emails: {e}"})


@mcp.tool()
def get_email(item_id: str, ctx: Context = None) -> str:
    """Get a single email with full body and details.

    Args:
        item_id: The Exchange ItemId of the email to retrieve.
    """
    try:
        client = _get_client(ctx)
        result = _get_item_details(client, item_id)
        return json.dumps(result)
    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to get email: {e}"})


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    importance: str = "Normal",
    is_html: bool = False,
    ctx: Context = None,
) -> str:
    """Send a new email.

    Args:
        to: Comma-separated list of recipient email addresses.
        subject: Email subject line.
        body: Email body text.
        cc: Comma-separated CC recipients (optional).
        bcc: Comma-separated BCC recipients (optional).
        importance: Email importance: Low, Normal, or High (default Normal).
        is_html: If True, body is treated as HTML. Otherwise plain text.
    """
    try:
        client = _get_client(ctx)

        to_recipients = _build_recipient_list(to)
        if not to_recipients:
            return json.dumps({"error": "At least one recipient is required."})

        message = {
            "__type": "Message:#Exchange",
            "Subject": subject,
            "Body": {
                "__type": "BodyContentType:#Exchange",
                "BodyType": "HTML" if is_html else "Text",
                "Value": body,
            },
            "Importance": importance,
            "ToRecipients": to_recipients,
        }

        if cc:
            message["CcRecipients"] = _build_recipient_list(cc)
        if bcc:
            message["BccRecipients"] = _build_recipient_list(bcc)

        payload = {
            "__type": "CreateItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2017_08_18",
            },
            "Body": {
                "__type": "CreateItemRequest:#Exchange",
                "Items": [message],
                "MessageDisposition": "SendAndSaveCopy",
            },
        }

        data = client.request("CreateItem", payload)

        for msg in client.extract_items(data):
            if msg.get("ResponseClass") == "Success":
                return json.dumps({"success": True, "message": "Email sent."})
            elif msg.get("ResponseClass") == "Error":
                return json.dumps(
                    {"error": msg.get("MessageText", "Failed to send email.")}
                )

        return json.dumps({"success": True, "message": "Email sent."})

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to send email: {e}"})


@mcp.tool()
def reply_email(
    item_id: str,
    body: str,
    reply_all: bool = False,
    ctx: Context = None,
) -> str:
    """Reply to an email.

    Args:
        item_id: The Exchange ItemId of the email to reply to.
        body: Reply body text.
        reply_all: If True, reply to all recipients. Otherwise reply to sender only.
    """
    try:
        client = _get_client(ctx)

        change_key = _get_change_key(client, item_id)
        if not change_key:
            return json.dumps({"error": "Could not resolve item ChangeKey."})

        item_type = "ReplyAllToItem:#Exchange" if reply_all else "ReplyToItem:#Exchange"

        reply_item = {
            "__type": item_type,
            "ReferenceItemId": {
                "__type": "ItemId:#Exchange",
                "Id": item_id,
                "ChangeKey": change_key,
            },
            "NewBodyContent": {
                "__type": "BodyContentType:#Exchange",
                "BodyType": "Text",
                "Value": body,
            },
        }

        payload = {
            "__type": "CreateItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2017_08_18",
            },
            "Body": {
                "__type": "CreateItemRequest:#Exchange",
                "Items": [reply_item],
                "MessageDisposition": "SendAndSaveCopy",
            },
        }

        data = client.request("CreateItem", payload)

        for msg in client.extract_items(data):
            if msg.get("ResponseClass") == "Error":
                return json.dumps(
                    {"error": msg.get("MessageText", "Failed to send reply.")}
                )

        return json.dumps({"success": True, "message": "Reply sent."})

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to reply: {e}"})


@mcp.tool()
def forward_email(
    item_id: str,
    to: str,
    body: str = "",
    ctx: Context = None,
) -> str:
    """Forward an email to other recipients.

    Args:
        item_id: The Exchange ItemId of the email to forward.
        to: Comma-separated list of recipient email addresses.
        body: Optional message to include above the forwarded content.
    """
    try:
        client = _get_client(ctx)

        to_recipients = _build_recipient_list(to)
        if not to_recipients:
            return json.dumps({"error": "At least one recipient is required."})

        change_key = _get_change_key(client, item_id)
        if not change_key:
            return json.dumps({"error": "Could not resolve item ChangeKey."})

        forward_item = {
            "__type": "ForwardItem:#Exchange",
            "ReferenceItemId": {
                "__type": "ItemId:#Exchange",
                "Id": item_id,
                "ChangeKey": change_key,
            },
            "ToRecipients": to_recipients,
        }

        if body:
            forward_item["NewBodyContent"] = {
                "__type": "BodyContentType:#Exchange",
                "BodyType": "Text",
                "Value": body,
            }

        payload = {
            "__type": "CreateItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2017_08_18",
            },
            "Body": {
                "__type": "CreateItemRequest:#Exchange",
                "Items": [forward_item],
                "MessageDisposition": "SendAndSaveCopy",
            },
        }

        data = client.request("CreateItem", payload)

        for msg in client.extract_items(data):
            if msg.get("ResponseClass") == "Error":
                return json.dumps(
                    {"error": msg.get("MessageText", "Failed to forward.")}
                )

        return json.dumps({"success": True, "message": "Email forwarded."})

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to forward email: {e}"})


@mcp.tool()
def mark_email_read(
    item_ids: list[str],
    is_read: bool = True,
    ctx: Context = None,
) -> str:
    """Mark one or more emails as read or unread.

    Args:
        item_ids: List of Exchange ItemIds to update.
        is_read: True to mark as read, False to mark as unread (default True).
    """
    try:
        client = _get_client(ctx)

        changes = []
        for iid in item_ids:
            change_key = _get_change_key(client, iid)
            item_id_dict = {"__type": "ItemId:#Exchange", "Id": iid}
            if change_key:
                item_id_dict["ChangeKey"] = change_key
            changes.append(
                {
                    "__type": "ItemChange:#Exchange",
                    "ItemId": item_id_dict,
                    "Updates": [
                        {
                            "__type": "SetItemField:#Exchange",
                            "Path": {
                                "__type": "PropertyUri:#Exchange",
                                "FieldURI": "IsRead",
                            },
                            "Item": {
                                "__type": "Message:#Exchange",
                                "IsRead": is_read,
                            },
                        }
                    ],
                }
            )

        payload = {
            "__type": "UpdateItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2017_08_18",
            },
            "Body": {
                "__type": "UpdateItemRequest:#Exchange",
                "ItemChanges": changes,
                "ConflictResolution": "AutoResolve",
                "MessageDisposition": "SaveOnly",
            },
        }

        data = client.request("UpdateItem", payload)

        errors = []
        for msg in client.extract_items(data):
            if msg.get("ResponseClass") == "Error":
                errors.append(msg.get("MessageText", "Unknown error"))

        if errors:
            return json.dumps({"error": "; ".join(errors)})

        status = "read" if is_read else "unread"
        return json.dumps(
            {
                "success": True,
                "message": f"Marked {len(item_ids)} email(s) as {status}.",
            }
        )

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to update emails: {e}"})


@mcp.tool()
def move_email(
    item_ids: list[str],
    target_folder: str,
    ctx: Context = None,
) -> str:
    """Move one or more emails to a different folder.

    Args:
        item_ids: List of Exchange ItemIds to move.
        target_folder: Destination folder name (e.g. Inbox, Sent, Deleted, or custom).
    """
    try:
        client = _get_client(ctx)

        folder_id = client.get_folder_id(target_folder)
        if not folder_id:
            return json.dumps({"error": f"Folder '{target_folder}' not found."})

        items = [
            {"__type": "ItemId:#Exchange", "Id": iid} for iid in item_ids
        ]

        payload = {
            "__type": "MoveItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2017_08_18",
            },
            "Body": {
                "__type": "MoveItemRequest:#Exchange",
                "ItemIds": items,
                "ToFolderId": {
                    "__type": "TargetFolderId:#Exchange",
                    "BaseFolderId": {
                        "__type": "FolderId:#Exchange",
                        "Id": folder_id,
                    },
                },
            },
        }

        data = client.request("MoveItem", payload)

        errors = []
        for msg in client.extract_items(data):
            if msg.get("ResponseClass") == "Error":
                errors.append(msg.get("MessageText", "Unknown error"))

        if errors:
            return json.dumps({"error": "; ".join(errors)})

        return json.dumps(
            {
                "success": True,
                "message": f"Moved {len(item_ids)} email(s) to '{target_folder}'.",
            }
        )

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to move emails: {e}"})


@mcp.tool()
def delete_email(
    item_ids: list[str],
    permanent: bool = False,
    ctx: Context = None,
) -> str:
    """Delete one or more emails.

    Args:
        item_ids: List of Exchange ItemIds to delete.
        permanent: If True, permanently delete (HardDelete). Otherwise move to Deleted Items.
    """
    try:
        client = _get_client(ctx)

        items = [
            {"__type": "ItemId:#Exchange", "Id": iid} for iid in item_ids
        ]

        payload = {
            "__type": "DeleteItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2017_08_18",
            },
            "Body": {
                "__type": "DeleteItemRequest:#Exchange",
                "ItemIds": items,
                "DeleteType": "HardDelete" if permanent else "MoveToDeletedItems",
            },
        }

        data = client.request("DeleteItem", payload)

        errors = []
        for msg in client.extract_items(data):
            if msg.get("ResponseClass") == "Error":
                errors.append(msg.get("MessageText", "Unknown error"))

        if errors:
            return json.dumps({"error": "; ".join(errors)})

        action = "permanently deleted" if permanent else "moved to Deleted Items"
        return json.dumps(
            {
                "success": True,
                "message": f"{len(item_ids)} email(s) {action}.",
            }
        )

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to delete emails: {e}"})


@mcp.tool()
def download_attachments(
    item_id: str,
    target_folder: str = "/tmp/attachments",
    ctx: Context = None,
) -> str:
    """Download all file attachments from an email to disk.

    Args:
        item_id: The Exchange ItemId of the email to download attachments from.
        target_folder: Local directory to save files (default /tmp/attachments).
    """
    import os

    try:
        client = _get_client(ctx)

        # Get email details to find attachment IDs
        details = _get_item_details(client, item_id)
        attachments = details.get("attachments", [])

        if not attachments:
            return json.dumps({"success": True, "downloaded": [], "count": 0,
                               "message": "No attachments found."})

        # Filter to non-inline file attachments with IDs
        file_attachments = [
            a for a in attachments
            if a.get("attachment_id") and not a.get("is_inline", False)
        ]

        if not file_attachments:
            return json.dumps({"success": True, "downloaded": [], "count": 0,
                               "message": "No downloadable file attachments."})

        os.makedirs(target_folder, exist_ok=True)

        downloaded = []
        errors = []
        used_names: set[str] = set()

        for att in file_attachments:
            try:
                content, filename, content_type = client.download_file(
                    att["attachment_id"]
                )

                # Sanitize filename
                filename = os.path.basename(filename)
                if not filename:
                    filename = att.get("name", "attachment") or "attachment"

                # Handle collisions
                base_name = filename
                name_part, _, ext_part = base_name.rpartition(".")
                if not name_part:
                    name_part = base_name
                    ext_part = ""

                counter = 1
                while filename.lower() in used_names:
                    if ext_part:
                        filename = f"{name_part}_{counter}.{ext_part}"
                    else:
                        filename = f"{name_part}_{counter}"
                    counter += 1

                used_names.add(filename.lower())

                filepath = os.path.join(target_folder, filename)
                with open(filepath, "wb") as f:
                    f.write(content)

                downloaded.append({
                    "name": filename,
                    "path": filepath,
                    "size": len(content),
                    "content_type": content_type,
                })
            except Exception as e:
                errors.append({
                    "name": att.get("name", "unknown"),
                    "error": str(e),
                })

        result = {
            "success": len(errors) == 0,
            "downloaded": downloaded,
            "count": len(downloaded),
        }
        if errors:
            result["errors"] = errors

        return json.dumps(result)

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to download attachments: {e}"})


@mcp.tool()
def get_email_links(
    item_id: str,
    ctx: Context = None,
) -> str:
    """Extract all hyperlinks from an email's HTML body.

    Args:
        item_id: The Exchange ItemId of the email to extract links from.
    """
    try:
        client = _get_client(ctx)

        # Fetch email with HTML body
        payload = {
            "__type": "GetItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2017_08_18",
            },
            "Body": {
                "__type": "GetItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                    "BodyType": "HTML",
                    "AdditionalProperties": [
                        {
                            "__type": "PropertyUri:#Exchange",
                            "FieldURI": "Subject",
                        },
                        {
                            "__type": "PropertyUri:#Exchange",
                            "FieldURI": "Body",
                        },
                    ],
                },
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": item_id}],
            },
        }

        data = client.request("GetItem", payload)

        subject = ""
        links = []

        for msg in client.extract_items(data):
            if "Items" not in msg:
                continue
            for item in msg["Items"]:
                subject = item.get("Subject", "")
                body_val = item.get("Body", {}).get("Value", "")
                links = extract_links_from_html(body_val)
                break

        return json.dumps({
            "item_id": item_id,
            "subject": subject,
            "links": links,
            "count": len(links),
        })

    except SessionExpiredError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": f"Failed to extract links: {e}"})
