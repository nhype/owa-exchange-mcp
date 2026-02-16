"""Calendar tools for the Exchange MCP server.

Provides MCP tools for calendar event retrieval, meeting creation,
update, cancellation, and meeting response management via OWA API.
"""

import html as html_mod
import json
import uuid
from datetime import datetime, timedelta

from mcp.server.fastmcp import Context

from exchange_mcp.server import mcp, AppContext
from exchange_mcp.owa_client import OWAClient
from exchange_mcp.utils import html_to_text, parse_iso_datetime, extract_links_from_html


def _get_client(ctx: Context) -> OWAClient:
    """Extract the OWAClient from the MCP lifespan context."""
    app_ctx: AppContext = ctx.request_context.lifespan_context
    return app_ctx.client


# ------------------------------------------------------------------
# Internal helpers
# ------------------------------------------------------------------


def _utc_to_local_str(dt_str: str) -> str:
    """Convert a UTC ISO timestamp to local Moscow time string.

    FindItem returns UTC (e.g. '2026-02-17T06:30:00Z'), while
    GetUserAvailability returns local Moscow time ('2026-02-17T09:30:00').
    This normalizes to the local format for key matching.
    Moscow is permanently UTC+3 (no DST since 2014).
    """
    if not dt_str:
        return dt_str
    if dt_str.endswith("Z"):
        try:
            utc_dt = datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%SZ")
            local_dt = utc_dt + timedelta(hours=3)
            return local_dt.strftime("%Y-%m-%dT%H:%M:%S")
        except ValueError:
            return dt_str.rstrip("Z")
    # No timezone suffix — already local time
    return dt_str


def _get_event_details(client: OWAClient, item_id: str) -> dict:
    """Get full event details (body, organizer, attendees) via GetItem."""
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
                "BaseShape": "AllProperties",
            },
            "ItemIds": [
                {"__type": "ItemId:#Exchange", "Id": item_id}
            ],
        },
    }

    data = client.request("GetItem", payload)
    result = {
        "organizer": "",
        "organizer_email": "",
        "location": "",
        "body": "",
        "attendees_required": [],
        "attendees_optional": [],
    }

    for msg in client.extract_items(data):
        if "Items" not in msg:
            continue
        for item in msg["Items"]:
            # Location
            result["location"] = item.get("Location", "")
            if not result["location"]:
                enhanced = item.get("EnhancedLocation", {})
                if enhanced:
                    result["location"] = enhanced.get("DisplayName", "")

            # Body
            body_data = item.get("Body", {})
            if body_data:
                body_text = body_data.get("Value", "")
                if body_data.get("BodyType") == "HTML":
                    body_text = html_to_text(body_text)
                result["body"] = body_text.strip()

            # Organizer with SMTP email
            organizer = item.get("Organizer", {}).get("Mailbox", {})
            if organizer:
                name = organizer.get("Name", "")
                addr = organizer.get("EmailAddress", "")
                if addr and not addr.startswith("/O="):
                    result["organizer"] = f"{name} <{addr}>" if name else addr
                    result["organizer_email"] = addr
                else:
                    result["organizer"] = name

            # Required attendees with SMTP emails
            for a in item.get("RequiredAttendees", []) or []:
                mailbox = a.get("Mailbox", {})
                name = mailbox.get("Name", "")
                addr = mailbox.get("EmailAddress", "")
                response = a.get("ResponseType", "")

                if name or addr:
                    if addr and not addr.startswith("/O="):
                        entry = f"{name} <{addr}>" if name else addr
                    else:
                        entry = name
                    if response and response not in ("Unknown", "Organizer"):
                        entry += f" [{response}]"
                    result["attendees_required"].append(entry)

            # Optional attendees with SMTP emails
            for a in item.get("OptionalAttendees", []) or []:
                mailbox = a.get("Mailbox", {})
                name = mailbox.get("Name", "")
                addr = mailbox.get("EmailAddress", "")
                response = a.get("ResponseType", "")

                if name or addr:
                    if addr and not addr.startswith("/O="):
                        entry = f"{name} <{addr}>" if name else addr
                    else:
                        entry = name
                    if response and response not in ("Unknown", "Organizer"):
                        entry += f" [{response}]"
                    result["attendees_optional"].append(entry)

            return result

    return result


def _get_full_event(client: OWAClient, item_id: str) -> dict:
    """Get full event details for update_meeting preservation."""
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
                "BaseShape": "AllProperties",
            },
            "ItemIds": [
                {"__type": "ItemId:#Exchange", "Id": item_id}
            ],
        },
    }

    data = client.request("GetItem", payload)
    for msg in client.extract_items(data):
        if "Items" not in msg:
            continue
        for item in msg["Items"]:
            result = {
                "subject": item.get("Subject", ""),
                "start": item.get("Start", ""),
                "end": item.get("End", ""),
                "is_all_day": item.get("IsAllDayEvent", False),
                "sensitivity": item.get("Sensitivity", "Normal"),
                "location": "",
                "body_html": "",
                "resolved_required": [],
                "resolved_optional": [],
            }

            # Location
            loc = item.get("Location", "")
            if not loc:
                enhanced = item.get("EnhancedLocation", {})
                if enhanced:
                    loc = enhanced.get("DisplayName", "")
            result["location"] = loc

            # Body (keep HTML)
            body_data = item.get("Body", {})
            if body_data:
                result["body_html"] = body_data.get("Value", "")

            # Required attendees as resolved dicts
            for a in item.get("RequiredAttendees", []) or []:
                mailbox = a.get("Mailbox", {})
                name = mailbox.get("Name", "")
                addr = mailbox.get("EmailAddress", "")
                if addr and not addr.startswith("/O="):
                    result["resolved_required"].append({
                        "Mailbox": {
                            "Name": name,
                            "EmailAddress": addr,
                            "RoutingType": "SMTP",
                        }
                    })

            # Optional attendees
            for a in item.get("OptionalAttendees", []) or []:
                mailbox = a.get("Mailbox", {})
                name = mailbox.get("Name", "")
                addr = mailbox.get("EmailAddress", "")
                if addr and not addr.startswith("/O="):
                    result["resolved_optional"].append({
                        "Mailbox": {
                            "Name": name,
                            "EmailAddress": addr,
                            "RoutingType": "SMTP",
                        }
                    })

            return result

    return {"error": "Meeting not found"}


def _resolve_attendee(client: OWAClient, email: str) -> dict:
    """Resolve an email to attendee details via ResolveNames.

    Uses V2017_08_18 RequestServerVersion to match the original
    create-meeting.py behaviour.
    """
    payload = {
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "V2017_08_18",
        },
        "Body": {
            "__type": "ResolveNamesRequest:#Exchange",
            "UnresolvedEntry": email,
            "ReturnFullContactData": True,
            "ContactDataShape": "Default",
        },
    }

    try:
        data = client.request("ResolveNames", payload)
        items = client.extract_items(data)
        if items and "ResolutionSet" in items[0]:
            resolutions = items[0]["ResolutionSet"].get("Resolutions", [])
            if resolutions:
                mailbox = resolutions[0].get("Mailbox", {})
                return {
                    "Mailbox": {
                        "Name": mailbox.get("Name", email),
                        "EmailAddress": mailbox.get("EmailAddress", email),
                        "RoutingType": "SMTP",
                    }
                }
    except Exception:
        pass

    # Fallback
    return {
        "Mailbox": {
            "Name": email,
            "EmailAddress": email,
            "RoutingType": "SMTP",
        }
    }


def _resolve_attendee_list(client: OWAClient, emails: list[str]) -> list[dict]:
    """Resolve a list of email strings into attendee dicts."""
    attendees = []
    for email in emails:
        email = email.strip()
        if email:
            attendees.append(_resolve_attendee(client, email))
    return attendees


def _build_html_body(description: str | None) -> str:
    """Build the HTML body for a calendar item, matching create-meeting.py."""
    body = (
        '<html><head><meta http-equiv="Content-Type" '
        'content="text/html; charset=UTF-8"></head><body dir="ltr">'
    )
    if description:
        desc_escaped = html_mod.escape(description).replace("\n", "<br>")
        body += (
            '<div style="font-size:12pt;color:#000000;'
            f'font-family:Calibri,Helvetica,sans-serif;">{desc_escaped}</div>'
        )
    else:
        body += (
            '<div style="font-size:12pt;color:#000000;'
            'font-family:Calibri,Helvetica,sans-serif;"><p><br></p></div>'
        )
    body += "</body></html>"
    return body


# ------------------------------------------------------------------
# Tool 1: get_calendar_events
# ------------------------------------------------------------------


def _get_expanded_events(
    client: OWAClient, start_date, end_date, chunk_days: int = 14
) -> list[dict]:
    """Get expanded calendar events via GetUserAvailability.

    Unlike FindItem (which returns only master items for recurring series),
    this returns every individual occurrence within the date range.

    Requires client.user_email to be set (done by the login tool).
    """
    if not client.user_email:
        return []

    expanded = []
    current = start_date

    while current < end_date:
        chunk_end = min(current + timedelta(days=chunk_days), end_date)

        payload = {
            '__type': 'GetUserAvailabilityJsonRequest:#Exchange',
            'Header': {
                '__type': 'JsonRequestHeaders:#Exchange',
                'RequestServerVersion': 'Exchange2013',
                'TimeZoneContext': {
                    '__type': 'TimeZoneContext:#Exchange',
                    'TimeZoneDefinition': {
                        '__type': 'TimeZoneDefinitionType:#Exchange',
                        'Id': 'Russian Standard Time',
                    },
                },
            },
            'Body': {
                '__type': 'GetUserAvailabilityRequest:#Exchange',
                'MailboxDataArray': [{
                    '__type': 'MailboxData:#Exchange',
                    'Email': {'__type': 'EmailAddress:#Exchange', 'Address': client.user_email},
                    'AttendeeType': 'Required',
                }],
                'FreeBusyViewOptions': {
                    '__type': 'FreeBusyViewOptions:#Exchange',
                    'TimeWindow': {
                        '__type': 'Duration:#Exchange',
                        'StartTime': f'{current}T00:00:00',
                        'EndTime': f'{chunk_end}T00:00:00',
                    },
                    'MergedFreeBusyIntervalInMinutes': 30,
                    'RequestedView': 'DetailedMerged',
                },
            },
        }

        try:
            data = client.request('GetUserAvailability', payload)
            body = data.get('Body', {})
            for fb_resp in body.get('FreeBusyResponseArray', []):
                fb_view = fb_resp.get('FreeBusyView', {})
                cal_events = fb_view.get('CalendarEventArray', {})
                items = (
                    cal_events.get('Items', [])
                    if isinstance(cal_events, dict)
                    else (cal_events if isinstance(cal_events, list) else [])
                )
                for event in items:
                    bt = event.get('BusyType', '')
                    details = event.get('CalendarEventDetails', {})
                    subject = details.get('Subject', '') if details else ''
                    location = details.get('Location', '') if details else ''
                    is_meeting = details.get('IsMeeting', False) if details else False
                    is_recurring = details.get('IsRecurring', False) if details else False

                    expanded.append({
                        'subject': subject or '(No subject)',
                        'start': event.get('StartTime', ''),
                        'end': event.get('EndTime', ''),
                        'busy_type': bt,
                        'location': location,
                        'is_meeting': is_meeting,
                        'is_recurring': is_recurring,
                    })
        except Exception:
            pass

        current = chunk_end

    return sorted(expanded, key=lambda x: x.get('start', ''))


@mcp.tool()
def get_calendar_events(
    start_date: str,
    end_date: str,
    include_body: bool = True,
    expand_recurring: bool = False,
    ctx: Context = None,
) -> str:
    """Get calendar events within a date range.

    Args:
        start_date: Start date in YYYY-MM-DD format.
        end_date: End date in YYYY-MM-DD format.
        include_body: If True, fetch full event details (organizer, attendees, body)
                      via GetItem for each event. Slower but more complete.
                      Ignored when expand_recurring=True.
        expand_recurring: If True, show every individual occurrence of recurring
                          meetings (via GetUserAvailability). This gives an accurate
                          count of all events but returns fewer fields per event
                          (no item_id, attendees, or body). Default False.

    Returns:
        JSON array of event objects with subject, start, end, location, attendees, etc.
    """
    client = _get_client(ctx)

    try:
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError as e:
        return json.dumps({"error": f"Invalid date format: {e}"})

    # --- Expanded mode: uses GetUserAvailability for accurate recurring counts ---
    if expand_recurring:
        expanded = _get_expanded_events(
            client, start_dt.date(), (end_dt + timedelta(days=1)).date()
        )
        events = []
        for ev in expanded:
            events.append({
                "subject": ev['subject'],
                "start": ev['start'],
                "end": ev['end'],
                "location": ev.get('location', ''),
                "busy_type": ev.get('busy_type', ''),
                "is_meeting": ev.get('is_meeting', False),
                "is_recurring": ev.get('is_recurring', False),
            })
        return json.dumps(events, ensure_ascii=False)

    # --- Default mode ---
    # Step 1: Get all events (including recurring) via GetUserAvailability
    expanded = _get_expanded_events(
        client, start_dt.date(), (end_dt + timedelta(days=1)).date()
    )

    # Step 2: Get events with item_ids via FindItem + CalendarView.
    # CalendarView restricts results to the date range and expands recurring
    # events into individual occurrences (each with its own ItemId).
    folder_id = client.get_folder_id("calendar")
    finditem_by_key: dict[str, dict] = {}  # "subject|start" -> item
    if folder_id:
        cv_start = start_dt.strftime("%Y-%m-%dT00:00:00")
        cv_end = (end_dt + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00")

        payload = {
            "__type": "FindItemJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "Exchange2013",
            },
            "Body": {
                "__type": "FindItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "AllProperties",
                },
                "ParentFolderIds": [
                    {"__type": "FolderId:#Exchange", "Id": folder_id}
                ],
                "Traversal": "Shallow",
                "CalendarView": {
                    "__type": "CalendarView:#Exchange",
                    "StartDate": cv_start,
                    "EndDate": cv_end,
                },
            },
        }

        data = client.request("FindItem", payload)

        all_items = []
        for msg in client.extract_items(data):
            if "RootFolder" in msg:
                all_items = msg["RootFolder"].get("Items", [])
                break

        # Index FindItem results by subject + normalized local start time.
        # FindItem returns UTC timestamps (e.g. "2026-02-17T06:30:00Z"),
        # while GetUserAvailability returns Moscow local time ("2026-02-17T09:30:00").
        # Normalize FindItem timestamps to local time for matching.
        for item in all_items:
            subject = item.get("Subject", "")
            start = item.get("Start", "")
            local_start = _utc_to_local_str(start)
            key = f"{subject}|{local_start}"
            finditem_by_key[key] = item

    # Step 3: Merge — use expanded list as the authoritative event list,
    # enrich with FindItem data (item_id, details) when available
    events = []
    for ev in expanded:
        subject = ev.get("subject", "(No subject)")
        start = ev.get("start", "")
        end = ev.get("end", "")
        is_recurring = ev.get("is_recurring", False)

        # Match expanded event (local time) with FindItem result (normalized to local)
        fi_item = finditem_by_key.get(f"{subject}|{start}")

        item_id = fi_item.get("ItemId", {}).get("Id", "") if fi_item else ""

        event = {
            "subject": subject,
            "start": fi_item.get("Start", start) if fi_item else start,
            "end": fi_item.get("End", end) if fi_item else end,
            "location": ev.get("location", ""),
            "is_all_day": fi_item.get("IsAllDayEvent", False) if fi_item else False,
            "is_cancelled": fi_item.get("IsCancelled", False) if fi_item else False,
            "is_meeting": ev.get("is_meeting", False),
            "is_recurring": is_recurring,
            "organizer": "",
            "my_response": fi_item.get("MyResponseType", "") if fi_item else "",
            "item_id": item_id,
            "body": "",
            "attendees_required": [],
            "attendees_optional": [],
        }

        # Get full details via GetItem if requested and item_id is available
        if include_body and item_id:
            details = _get_event_details(client, item_id)
            event["organizer"] = details["organizer"]
            event["location"] = details["location"] or event["location"]
            event["body"] = details["body"]
            event["attendees_required"] = details["attendees_required"]
            event["attendees_optional"] = details["attendees_optional"]

        events.append(event)

    return json.dumps(events, ensure_ascii=False)


# ------------------------------------------------------------------
# Tool 2: create_meeting
# ------------------------------------------------------------------


@mcp.tool()
def create_meeting(
    subject: str,
    date: str,
    start_time: str,
    duration_minutes: int = 30,
    required_attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None,
    location: str | None = None,
    description: str | None = None,
    is_all_day: bool = False,
    reminder_minutes: int = 15,
    importance: str = "Normal",
    sensitivity: str = "Normal",
    ctx: Context = None,
) -> str:
    """Create a new calendar meeting.

    Args:
        subject: Meeting subject/topic.
        date: Meeting date in YYYY-MM-DD format.
        start_time: Start time in HH:MM format.
        duration_minutes: Duration in minutes (default 30).
        required_attendees: List of email addresses for required attendees.
        optional_attendees: List of email addresses for optional attendees.
        location: Location or video link.
        description: Meeting description/body text.
        is_all_day: Whether this is an all-day event.
        reminder_minutes: Minutes before start for reminder (default 15).
        importance: Importance level: Low, Normal, or High.
        sensitivity: Sensitivity: Normal, Personal, Private, or Confidential.

    Returns:
        JSON object with creation result including item_id on success.
    """
    client = _get_client(ctx)

    # Parse date and time
    try:
        start_dt = datetime.strptime(f"{date} {start_time}", "%Y-%m-%d %H:%M")
        end_dt = start_dt + timedelta(minutes=duration_minutes)
    except ValueError as e:
        return json.dumps({"error": f"Invalid date/time: {e}"})

    # Resolve attendees
    resolved_required = _resolve_attendee_list(client, required_attendees or [])
    resolved_optional = _resolve_attendee_list(client, optional_attendees or [])

    # Build HTML body
    html_body = _build_html_body(description)

    # Build location object
    location_obj = {
        "__type": "EnhancedLocation:#Exchange",
        "Annotation": "",
        "DisplayName": location or "",
        "PostalAddress": {
            "__type": "PersonaPostalAddress:#Exchange",
            "Type": "Business",
            "LocationSource": "None",
        },
    }

    # Build calendar item
    calendar_item = {
        "__type": "CalendarItem:#Exchange",
        "ClientSeriesId": str(uuid.uuid4()),
        "Subject": subject,
        "Body": {
            "__type": "BodyContentType:#Exchange",
            "BodyType": "HTML",
            "Value": html_body,
        },
        "Sensitivity": sensitivity,
        "ReminderIsSet": True,
        "ReminderMinutesBeforeStart": reminder_minutes,
        "IsResponseRequested": True,
        "DoNotForwardMeeting": False,
        "IsAllDayEvent": is_all_day,
        "Start": start_dt.strftime("%Y-%m-%dT%H:%M:%S.000"),
        "End": end_dt.strftime("%Y-%m-%dT%H:%M:%S.000"),
        "FreeBusyType": "Busy",
        "Location": location_obj,
        "unfoldedIndex": 0,
    }

    if importance != "Normal":
        calendar_item["Importance"] = importance

    if resolved_required:
        calendar_item["RequiredAttendees"] = resolved_required
    if resolved_optional:
        calendar_item["OptionalAttendees"] = resolved_optional

    # Build request - uses CreateCalendarEvent action and V2017_08_18
    payload = {
        "__type": "CreateItemJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "V2017_08_18",
            "TimeZoneContext": {
                "__type": "TimeZoneContext:#Exchange",
                "TimeZoneDefinition": {
                    "__type": "TimeZoneDefinitionType:#Exchange",
                    "Id": "Russian Standard Time",
                },
            },
        },
        "Body": {
            "__type": "CreateItemRequest:#Exchange",
            "Items": [calendar_item],
            "ClientSupportsIrm": True,
            "SavedItemFolderId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": {
                    "__type": "DistinguishedFolderId:#Exchange",
                    "Id": "calendar",
                },
            },
        },
    }

    # Send invitations if there are attendees
    if resolved_required or resolved_optional:
        payload["Body"]["SendMeetingInvitations"] = "SendToAllAndSaveCopy"

    data = client.request("CreateCalendarEvent", payload)

    # Check for top-level error
    body = data.get("Body", {})
    if "ErrorCode" in body:
        return json.dumps({"error": body.get("FaultMessage", "Unknown error")})

    # Check response messages
    items = body.get("ResponseMessages", {}).get("Items", [])
    if items:
        item = items[0]
        if item.get("ResponseClass") == "Success":
            result = {
                "success": True,
                "subject": subject,
                "date": date,
                "start_time": start_time,
                "end_time": end_dt.strftime("%H:%M"),
                "duration_minutes": duration_minutes,
            }
            # Extract created item ID if available
            created_items = item.get("Items", [])
            if created_items:
                item_id = created_items[0].get("ItemId", {})
                if item_id:
                    result["item_id"] = item_id.get("Id", "")
                    result["change_key"] = item_id.get("ChangeKey", "")
            if location:
                result["location"] = location
            if resolved_required:
                result["required_attendees"] = [
                    a["Mailbox"]["EmailAddress"] for a in resolved_required
                ]
            if resolved_optional:
                result["optional_attendees"] = [
                    a["Mailbox"]["EmailAddress"] for a in resolved_optional
                ]
            return json.dumps(result, ensure_ascii=False)
        else:
            return json.dumps({
                "error": item.get("MessageText", "Unknown error"),
                "response_code": item.get("ResponseCode", ""),
            })

    return json.dumps({"success": True, "subject": subject, "note": "No confirmation details"})


# ------------------------------------------------------------------
# Tool 3: update_meeting
# ------------------------------------------------------------------


@mcp.tool()
def update_meeting(
    item_id: str,
    subject: str | None = None,
    date: str | None = None,
    start_time: str | None = None,
    duration_minutes: int | None = None,
    location: str | None = None,
    description: str | None = None,
    required_attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None,
    change_key: str = "",
    ctx: Context = None,
) -> str:
    """Update an existing calendar meeting.

    Internally cancels the old meeting and creates a new one with
    updated fields, because OWA's JSON API does not support UpdateItem
    for calendar items reliably. Unchanged fields are preserved from
    the original meeting.

    Args:
        item_id: The ItemId of the meeting to update (from get_calendar_events).
        subject: New subject (omit to keep original).
        date: New date in YYYY-MM-DD format (omit to keep original).
        start_time: New start time in HH:MM format (omit to keep original).
        duration_minutes: New duration in minutes (omit to keep original).
        location: New location (omit to keep original).
        description: New description/body text (omit to keep original).
        required_attendees: Email addresses for required attendees.
            Replaces existing list. Omit to keep original attendees.
        optional_attendees: Email addresses for optional attendees.
            Replaces existing list. Omit to keep original attendees.
        change_key: Ignored (kept for backward compatibility).

    Returns:
        JSON object with update result including new item_id.
    """
    client = _get_client(ctx)

    # Step 1: Get the original meeting details
    try:
        orig = _get_full_event(client, item_id)
    except Exception as e:
        return json.dumps({"error": f"Could not fetch original meeting: {e}"})

    if "error" in orig:
        return json.dumps(orig)

    # Step 2: Merge original values with updates
    new_subject = subject if subject is not None else orig.get("subject", "")

    # Parse original start/end for date and time defaults
    orig_start_str = orig.get("start", "")
    orig_end_str = orig.get("end", "")
    try:
        orig_start = datetime.fromisoformat(orig_start_str.replace("Z", "+00:00")).replace(tzinfo=None)
        orig_end = datetime.fromisoformat(orig_end_str.replace("Z", "+00:00")).replace(tzinfo=None)
        orig_duration = int((orig_end - orig_start).total_seconds() / 60)
    except (ValueError, AttributeError):
        orig_start = None
        orig_end = None
        orig_duration = 30

    if date is not None and start_time is not None:
        new_start = datetime.strptime(f"{date} {start_time}", "%Y-%m-%d %H:%M")
        dur = duration_minutes if duration_minutes is not None else orig_duration
        new_end = new_start + timedelta(minutes=dur)
    elif date is not None and orig_start is not None:
        new_start = datetime.strptime(date, "%Y-%m-%d").replace(
            hour=orig_start.hour, minute=orig_start.minute
        )
        dur = duration_minutes if duration_minutes is not None else orig_duration
        new_end = new_start + timedelta(minutes=dur)
    elif start_time is not None and orig_start is not None:
        parts = start_time.split(":")
        new_start = orig_start.replace(hour=int(parts[0]), minute=int(parts[1]))
        dur = duration_minutes if duration_minutes is not None else orig_duration
        new_end = new_start + timedelta(minutes=dur)
    elif duration_minutes is not None and orig_start is not None:
        new_start = orig_start
        new_end = new_start + timedelta(minutes=duration_minutes)
    elif orig_start is not None:
        new_start = orig_start
        new_end = orig_end
    else:
        return json.dumps({"error": "Cannot determine meeting time. Provide date and start_time."})

    new_location = location if location is not None else orig.get("location", "")

    # Resolve attendees
    if required_attendees is not None:
        resolved_required = _resolve_attendee_list(client, required_attendees)
    else:
        resolved_required = orig.get("resolved_required", [])

    if optional_attendees is not None:
        resolved_optional = _resolve_attendee_list(client, optional_attendees)
    else:
        resolved_optional = orig.get("resolved_optional", [])

    # Step 3: Cancel the original meeting
    cancel_payload = {
        "__type": "DeleteItemJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "DeleteItemRequest:#Exchange",
            "ItemIds": [
                {"__type": "ItemId:#Exchange", "Id": item_id}
            ],
            "DeleteType": "MoveToDeletedItems",
            "SendMeetingCancellations": "SendToAllAndSaveCopy",
            "SuppressReadReceipts": True,
        },
    }

    try:
        client.request("DeleteItem", cancel_payload)
    except Exception as e:
        return json.dumps({"error": f"Failed to cancel original meeting: {e}"})

    # Step 4: Create the new meeting
    new_body = description if description is not None else orig.get("body_html", "")
    if description is not None:
        new_body = _build_html_body(description)

    location_obj = {
        "__type": "EnhancedLocation:#Exchange",
        "Annotation": "",
        "DisplayName": new_location,
        "PostalAddress": {
            "__type": "PersonaPostalAddress:#Exchange",
            "Type": "Business",
            "LocationSource": "None",
        },
    }

    calendar_item = {
        "__type": "CalendarItem:#Exchange",
        "ClientSeriesId": str(uuid.uuid4()),
        "Subject": new_subject,
        "Body": {
            "__type": "BodyContentType:#Exchange",
            "BodyType": "HTML",
            "Value": new_body,
        },
        "Sensitivity": orig.get("sensitivity", "Normal"),
        "ReminderIsSet": True,
        "ReminderMinutesBeforeStart": 15,
        "IsResponseRequested": True,
        "DoNotForwardMeeting": False,
        "IsAllDayEvent": orig.get("is_all_day", False),
        "Start": new_start.strftime("%Y-%m-%dT%H:%M:%S.000"),
        "End": new_end.strftime("%Y-%m-%dT%H:%M:%S.000"),
        "FreeBusyType": "Busy",
        "Location": location_obj,
        "unfoldedIndex": 0,
    }

    if resolved_required:
        calendar_item["RequiredAttendees"] = resolved_required
    if resolved_optional:
        calendar_item["OptionalAttendees"] = resolved_optional

    create_payload = {
        "__type": "CreateItemJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "V2017_08_18",
            "TimeZoneContext": {
                "__type": "TimeZoneContext:#Exchange",
                "TimeZoneDefinition": {
                    "__type": "TimeZoneDefinitionType:#Exchange",
                    "Id": "Russian Standard Time",
                },
            },
        },
        "Body": {
            "__type": "CreateItemRequest:#Exchange",
            "Items": [calendar_item],
            "ClientSupportsIrm": True,
            "SavedItemFolderId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": {
                    "__type": "DistinguishedFolderId:#Exchange",
                    "Id": "calendar",
                },
            },
        },
    }

    if resolved_required or resolved_optional:
        create_payload["Body"]["SendMeetingInvitations"] = "SendToAllAndSaveCopy"

    try:
        data = client.request("CreateCalendarEvent", create_payload)
    except Exception as e:
        return json.dumps({"error": f"Original cancelled but failed to create new: {e}"})

    body = data.get("Body", {})
    if "ErrorCode" in body:
        return json.dumps({"error": body.get("FaultMessage", "Unknown error")})

    resp_items = body.get("ResponseMessages", {}).get("Items", [])
    if resp_items and resp_items[0].get("ResponseClass") == "Success":
        result = {
            "success": True,
            "subject": new_subject,
            "start": new_start.strftime("%Y-%m-%d %H:%M"),
            "end": new_end.strftime("%Y-%m-%d %H:%M"),
            "duration_minutes": int((new_end - new_start).total_seconds() / 60),
        }
        created_items = resp_items[0].get("Items", [])
        if created_items:
            new_item_id = created_items[0].get("ItemId", {})
            if new_item_id:
                result["item_id"] = new_item_id.get("Id", "")
                result["change_key"] = new_item_id.get("ChangeKey", "")
        return json.dumps(result, ensure_ascii=False)

    if resp_items:
        return json.dumps({
            "error": resp_items[0].get("MessageText", "Unknown error"),
            "response_code": resp_items[0].get("ResponseCode", ""),
        })

    return json.dumps({"success": True, "subject": new_subject, "note": "No confirmation details"})


# ------------------------------------------------------------------
# Tool 4: cancel_meeting
# ------------------------------------------------------------------


@mcp.tool()
def cancel_meeting(
    item_id: str,
    message: str | None = None,
    ctx: Context = None,
) -> str:
    """Cancel (delete) a calendar meeting and notify attendees.

    Args:
        item_id: The ItemId of the meeting to cancel.
        message: Optional cancellation message to attendees.

    Returns:
        JSON object with cancellation result.
    """
    client = _get_client(ctx)

    payload = {
        "__type": "DeleteItemJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "DeleteItemRequest:#Exchange",
            "ItemIds": [
                {"__type": "ItemId:#Exchange", "Id": item_id}
            ],
            "DeleteType": "MoveToDeletedItems",
            "SendMeetingCancellations": "SendToAllAndSaveCopy",
            "SuppressReadReceipts": True,
        },
    }

    data = client.request("DeleteItem", payload)

    items = client.extract_items(data)
    if items:
        item = items[0]
        if item.get("ResponseClass") == "Success":
            return json.dumps({"success": True, "message": "Meeting cancelled"})
        else:
            return json.dumps({
                "error": item.get("MessageText", "Unknown error"),
                "response_code": item.get("ResponseCode", ""),
            })

    # DeleteItem may return empty on success
    body = data.get("Body", {})
    if "ErrorCode" in body:
        return json.dumps({"error": body.get("FaultMessage", "Unknown error")})

    return json.dumps({"success": True, "message": "Meeting cancelled"})


# ------------------------------------------------------------------
# Tool 5: respond_to_meeting
# ------------------------------------------------------------------


@mcp.tool()
def respond_to_meeting(
    item_id: str,
    response: str,
    message: str | None = None,
    ctx: Context = None,
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Args:
        item_id: The ItemId of the meeting to respond to.
        response: Response type: "Accept", "Decline", or "Tentative".
        message: Optional message to include with the response.

    Returns:
        JSON object with response result.
    """
    client = _get_client(ctx)

    # Map response to the correct __type
    response_types = {
        "Accept": "AcceptItem:#Exchange",
        "Decline": "DeclineItem:#Exchange",
        "Tentative": "TentativelyAcceptItem:#Exchange",
    }

    response_type = response_types.get(response)
    if not response_type:
        return json.dumps({
            "error": f"Invalid response: {response}. Must be Accept, Decline, or Tentative."
        })

    response_item = {
        "__type": response_type,
        "ReferenceItemId": {
            "__type": "ItemId:#Exchange",
            "Id": item_id,
        },
    }

    if message:
        response_item["Body"] = {
            "__type": "BodyContentType:#Exchange",
            "BodyType": "Text",
            "Value": message,
        }

    payload = {
        "__type": "CreateItemJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "CreateItemRequest:#Exchange",
            "Items": [response_item],
            "MessageDisposition": "SendAndSaveCopy",
        },
    }

    data = client.request("CreateItem", payload)

    items = client.extract_items(data)
    if items:
        item = items[0]
        if item.get("ResponseClass") == "Success":
            return json.dumps({
                "success": True,
                "response": response,
                "message": f"Meeting {response.lower()}ed",
            })
        else:
            return json.dumps({
                "error": item.get("MessageText", "Unknown error"),
                "response_code": item.get("ResponseCode", ""),
            })

    return json.dumps({"error": "No response from server"})


# ------------------------------------------------------------------
# Tool 6: download_event_attachments
# ------------------------------------------------------------------


@mcp.tool()
def download_event_attachments(
    item_id: str,
    target_folder: str = "/tmp/attachments",
    ctx: Context = None,
) -> str:
    """Download all file attachments from a calendar event to disk.

    Args:
        item_id: The Exchange ItemId of the calendar event.
        target_folder: Local directory to save files (default /tmp/attachments).
    """
    import os

    try:
        client = _get_client(ctx)

        # Get full event details (AllProperties includes Attachments)
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
                    "BaseShape": "AllProperties",
                },
                "ItemIds": [
                    {"__type": "ItemId:#Exchange", "Id": item_id}
                ],
            },
        }

        data = client.request("GetItem", payload)

        # Extract attachments from the item
        attachments = []
        for msg in client.extract_items(data):
            if "Items" not in msg:
                continue
            for item in msg["Items"]:
                for att in item.get("Attachments", []):
                    attachments.append({
                        "name": att.get("Name", ""),
                        "size": att.get("Size", 0),
                        "content_type": att.get("ContentType", ""),
                        "attachment_id": att.get("AttachmentId", {}).get("Id", ""),
                        "is_inline": att.get("IsInline", False),
                    })
                break

        if not attachments:
            return json.dumps({"success": True, "downloaded": [], "count": 0,
                               "message": "No attachments found on this event."})

        # Filter to non-inline file attachments
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

    except Exception as e:
        return json.dumps({"error": f"Failed to download event attachments: {e}"})


# ------------------------------------------------------------------
# Tool 7: get_event_links
# ------------------------------------------------------------------


@mcp.tool()
def get_event_links(
    item_id: str,
    ctx: Context = None,
) -> str:
    """Extract all hyperlinks from a calendar event's HTML description.

    Args:
        item_id: The Exchange ItemId of the calendar event to extract links from.
    """
    try:
        client = _get_client(ctx)

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

    except Exception as e:
        return json.dumps({"error": f"Failed to extract event links: {e}"})
