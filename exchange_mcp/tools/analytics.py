"""Analytics tools for the Exchange MCP server.

Provides meeting statistics and connection matrix analysis
using GetUserAvailability and GetItem APIs.
"""

import json
from collections import Counter, defaultdict
from datetime import datetime, timedelta, date

from mcp.server.fastmcp import Context

from exchange_mcp.server import mcp, AppContext
from exchange_mcp.owa_client import OWAClient


def _get_client(ctx: Context) -> OWAClient:
    """Extract the OWAClient from the MCP lifespan context."""
    app_ctx: AppContext = ctx.request_context.lifespan_context
    return app_ctx.client


# ------------------------------------------------------------------
# Internal: resolve names to emails
# ------------------------------------------------------------------

def _resolve_to_email(client: OWAClient, name: str) -> tuple[str, str]:
    """Resolve a name/email to (display_name, email). Returns ('','') on failure."""
    resolutions = client.resolve_names(name)
    if resolutions:
        mb = resolutions[0].get("Mailbox", {})
        return mb.get("Name", name), mb.get("EmailAddress", "")
    return "", ""


# ------------------------------------------------------------------
# Internal: query GetUserAvailability in chunks
# ------------------------------------------------------------------

def _get_availability_events(
    client: OWAClient,
    emails: list[str],
    start: date,
    end: date,
    batch_size: int = 5,
    chunk_days: int = 14,
) -> dict[str, list[dict]]:
    """Query GetUserAvailability for multiple people across a date range.

    Returns dict mapping email -> list of calendar event dicts with
    keys: subject, start_date, busy_type.
    """
    results: dict[str, list[dict]] = {email: [] for email in emails}

    # Batch people
    email_batches = [emails[i:i+batch_size] for i in range(0, len(emails), batch_size)]

    for batch in email_batches:
        current = start
        while current < end:
            chunk_end = min(current + timedelta(days=chunk_days), end)

            mailbox_data = [{
                '__type': 'MailboxData:#Exchange',
                'Email': {
                    '__type': 'EmailAddress:#Exchange',
                    'Address': email,
                },
                'AttendeeType': 'Required',
            } for email in batch]

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
                    'MailboxDataArray': mailbox_data,
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
                for i, fb_resp in enumerate(body.get('FreeBusyResponseArray', [])):
                    if i >= len(batch):
                        break
                    email = batch[i]
                    fb_view = fb_resp.get('FreeBusyView', {})
                    cal_events = fb_view.get('CalendarEventArray', {})
                    if isinstance(cal_events, dict):
                        items = cal_events.get('Items', [])
                    elif isinstance(cal_events, list):
                        items = cal_events
                    else:
                        items = []
                    for event in items:
                        s = event.get('StartTime', '')
                        bt = event.get('BusyType', '')
                        if s and bt != 'Free':
                            details = event.get('CalendarEventDetails', {})
                            subject = details.get('Subject', '') if details else ''
                            results[email].append({
                                'subject': subject,
                                'start_date': s[:10],
                                'busy_type': bt,
                            })
            except Exception:
                pass

            current = chunk_end

    return results


# ------------------------------------------------------------------
# Tool: get_meeting_stats
# ------------------------------------------------------------------

@mcp.tool()
def get_meeting_stats(
    people: str,
    start_date: str,
    end_date: str,
    ctx: Context = None,
) -> str:
    """Get meeting count statistics for one or more people over a date range.

    Uses GetUserAvailability to count calendar events (including
    expanded recurring instances) for each person.

    Args:
        people: Comma-separated names or email addresses to analyze.
        start_date: Start date in YYYY-MM-DD format.
        end_date: End date in YYYY-MM-DD format.

    Returns:
        JSON object with per-person stats sorted by meeting count:
        total_meetings, meetings_per_workday, days_with_meetings.
    """
    client = _get_client(ctx)

    try:
        sd = datetime.strptime(start_date, '%Y-%m-%d').date()
        ed = datetime.strptime(end_date, '%Y-%m-%d').date()
    except ValueError as e:
        return json.dumps({"error": f"Invalid date format: {e}"})

    # Resolve names
    name_list = [n.strip() for n in people.split(',') if n.strip()]
    if not name_list:
        return json.dumps({"error": "No people specified."})

    resolved = []  # (display_name, email)
    for name in name_list:
        display, email = _resolve_to_email(client, name)
        if email:
            resolved.append((display, email))
        else:
            resolved.append((name, ""))

    emails = [email for _, email in resolved if email]
    if not emails:
        return json.dumps({"error": "Could not resolve any names to email addresses."})

    # Query availability
    avail = _get_availability_events(client, emails, sd, ed)

    # Count working days
    workdays = 0
    d = sd
    while d < ed:
        if d.weekday() < 5:
            workdays += 1
        d += timedelta(days=1)
    workdays = max(workdays, 1)

    # Build stats
    stats = []
    for display, email in resolved:
        if not email or email not in avail:
            stats.append({
                "name": display,
                "email": email or "not_found",
                "total_meetings": 0,
                "meetings_per_workday": 0,
                "days_with_meetings": 0,
                "workdays": workdays,
            })
            continue

        events = avail[email]
        total = len(events)
        unique_days = len(set(ev['start_date'] for ev in events))
        avg = round(total / workdays, 1)

        stats.append({
            "name": display,
            "email": email,
            "total_meetings": total,
            "meetings_per_workday": avg,
            "days_with_meetings": unique_days,
            "workdays": workdays,
        })

    # Sort by total descending
    stats.sort(key=lambda x: x["total_meetings"], reverse=True)

    return json.dumps({
        "period": {"start": start_date, "end": end_date, "workdays": workdays},
        "stats": stats,
    }, ensure_ascii=False)


# ------------------------------------------------------------------
# Tool: get_meeting_contacts
# ------------------------------------------------------------------

@mcp.tool()
def get_meeting_contacts(
    start_date: str,
    end_date: str,
    top_n: int = 30,
    ctx: Context = None,
) -> str:
    """Build a connection matrix: who you meet with most often.

    Analyzes your calendar over a date range, accounting for recurring
    meetings. For each contact found in your meetings, returns the
    weighted count of shared meetings.

    Args:
        start_date: Start date in YYYY-MM-DD format.
        end_date: End date in YYYY-MM-DD format.
        top_n: Number of top contacts to return. Default 30.

    Returns:
        JSON object with total_meetings, unique_contacts, and a ranked
        contacts array of {name, email, meetings} objects.
    """
    client = _get_client(ctx)

    try:
        sd = datetime.strptime(start_date, '%Y-%m-%d').date()
        ed = datetime.strptime(end_date, '%Y-%m-%d').date()
    except ValueError as e:
        return json.dumps({"error": f"Invalid date format: {e}"})

    # Step 1: Get own expanded events via GetUserAvailability
    # to count subject occurrences (handles recurring meetings)
    own_email = client.user_email.lower() if client.user_email else ""
    if not own_email:
        return json.dumps({"error": "User email not available. Call the login tool first."})

    folder_id = client.get_folder_id("calendar")
    if not folder_id:
        return json.dumps({"error": "Could not find calendar folder."})

    # Get expanded event subjects with occurrence counts
    avail_result = _get_availability_events(client, [client.user_email], sd, ed)
    expanded_events = avail_result.get(client.user_email, [])

    subject_counts = Counter(ev['subject'] for ev in expanded_events)
    total_expanded = len(expanded_events)

    # Step 2: Find master calendar items for each unique subject
    # Scan entire calendar (descending) to match subjects
    subject_to_id: dict[str, str] = {}
    offset = 0

    while len(subject_to_id) < len(subject_counts):
        payload = {
            '__type': 'FindItemJsonRequest:#Exchange',
            'Header': {
                '__type': 'JsonRequestHeaders:#Exchange',
                'RequestServerVersion': 'Exchange2013',
            },
            'Body': {
                '__type': 'FindItemRequest:#Exchange',
                'ItemShape': {
                    '__type': 'ItemResponseShape:#Exchange',
                    'BaseShape': 'Default',
                },
                'ParentFolderIds': [
                    {'__type': 'FolderId:#Exchange', 'Id': folder_id}
                ],
                'Traversal': 'Shallow',
                'Paging': {
                    '__type': 'IndexedPageView:#Exchange',
                    'BasePoint': 'Beginning',
                    'Offset': offset,
                    'MaxEntriesReturned': 200,
                },
                'SortOrder': [{
                    '__type': 'SortResults:#Exchange',
                    'Order': 'Descending',
                    'Path': {
                        '__type': 'PropertyUri:#Exchange',
                        'FieldURI': 'Start',
                    },
                }],
            },
        }

        data = client.request('FindItem', payload)
        items = client.extract_items(data)
        if not items:
            break
        folder_items = items[0].get('RootFolder', {}).get('Items', [])
        is_last = items[0].get('RootFolder', {}).get('IncludesLastItemInRange', True)
        if not folder_items:
            break

        for item in folder_items:
            subj = item.get('Subject', '')
            if subj and subj in subject_counts and subj not in subject_to_id:
                subject_to_id[subj] = item.get('ItemId', {}).get('Id', '')

        if is_last:
            break
        offset += 200
        if offset > 3000:
            break

    # Step 3: GetItem for each master to get attendees
    unique_ids = list(set(subject_to_id.values()))
    id_to_attendees: dict[str, set] = {}

    for bi in range(0, len(unique_ids), 10):
        batch = unique_ids[bi:bi+10]
        item_id_list = [{"__type": "ItemId:#Exchange", "Id": iid} for iid in batch]

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
                "ItemIds": item_id_list,
            },
        }

        try:
            data = client.request("GetItem", payload)
            for msg in client.extract_items(data):
                if "Items" not in msg:
                    continue
                for item in msg["Items"]:
                    item_id = item.get("ItemId", {}).get("Id", "")
                    attendees = set()

                    for att_list in [item.get("RequiredAttendees", []) or [],
                                     item.get("OptionalAttendees", []) or []]:
                        for a in att_list:
                            mailbox = a.get("Mailbox", {})
                            name = mailbox.get("Name", "")
                            addr = mailbox.get("EmailAddress", "")
                            if not name or not addr or addr.startswith("/O="):
                                continue
                            attendees.add((name, addr.lower()))

                    organizer = item.get("Organizer", {}).get("Mailbox", {})
                    if organizer:
                        org_addr = organizer.get("EmailAddress", "")
                        org_name = organizer.get("Name", "")
                        if org_addr and not org_addr.startswith("/O="):
                            attendees.add((org_name, org_addr.lower()))

                    id_to_attendees[item_id] = attendees
        except Exception:
            pass

    # Step 4: Build weighted contacts, excluding self
    contacts = Counter()
    for subject, item_id in subject_to_id.items():
        weight = subject_counts.get(subject, 1)
        attendees = id_to_attendees.get(item_id, set())
        for name, email in attendees:
            if own_email and email == own_email:
                continue
            contacts[(name, email)] += weight

    result_contacts = []
    for (name, email), count in contacts.most_common(top_n):
        result_contacts.append({
            "name": name,
            "email": email,
            "meetings": count,
        })

    return json.dumps({
        "period": {"start": start_date, "end": end_date},
        "total_meetings": total_expanded,
        "unique_contacts": len(contacts),
        "contacts": result_contacts,
    }, ensure_ascii=False)
