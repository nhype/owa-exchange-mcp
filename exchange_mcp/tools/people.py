"""People / directory search tools for the Exchange MCP server.

Ports the find-person.py logic into an MCP tool using OWAClient.
"""

import json

from mcp.server.fastmcp import Context

from exchange_mcp.server import mcp, AppContext
from exchange_mcp.owa_client import OWAClient


def _get_client(ctx: Context) -> OWAClient:
    """Extract the OWAClient from the MCP lifespan context."""
    app_ctx: AppContext = ctx.request_context.lifespan_context
    return app_ctx.client


def _parse_person(resolution: dict) -> dict:
    """Parse person data from a ResolveNames resolution entry.

    Preserves the exact logic from find-person.py parse_person().
    """
    mailbox = resolution.get("Mailbox", {})
    contact = resolution.get("Contact", {})

    person = {
        "name": mailbox.get("Name", contact.get("DisplayName", "")),
        "email": mailbox.get("EmailAddress", ""),
        "type": mailbox.get("MailboxType", ""),
        "first_name": contact.get("GivenName", ""),
        "last_name": contact.get("Surname", ""),
        "job_title": contact.get("JobTitle", ""),
        "department": contact.get("Department", ""),
        "company": contact.get("CompanyName", ""),
        "office": contact.get("OfficeLocation", ""),
        "manager": "",
        "manager_email": "",
        "phones": {},
        "address": {},
        "direct_reports": [],
        "alias": contact.get("Alias", ""),
    }

    # Phone numbers
    for phone in contact.get("PhoneNumbers", []):
        key = phone.get("Key", "")
        number = phone.get("PhoneNumber", "")
        if number:
            person["phones"][key] = number

    # Physical address
    for addr in contact.get("PhysicalAddresses", []):
        if addr.get("Key") == "Business":
            parts = []
            if addr.get("Street"):
                parts.append(addr["Street"])
            if addr.get("City"):
                parts.append(addr["City"])
            if addr.get("PostalCode"):
                parts.append(addr["PostalCode"])
            if addr.get("CountryOrRegion"):
                parts.append(addr["CountryOrRegion"])
            if parts:
                person["address"] = {
                    "street": addr.get("Street", ""),
                    "city": addr.get("City", ""),
                    "postal_code": addr.get("PostalCode", ""),
                    "country": addr.get("CountryOrRegion", ""),
                    "full": ", ".join(parts),
                }

    # Manager
    manager_data = contact.get("ManagerMailbox", {}).get("Mailbox", {})
    if manager_data:
        person["manager"] = manager_data.get("Name", "")
        person["manager_email"] = manager_data.get("EmailAddress", "")
    elif contact.get("Manager"):
        person["manager"] = contact.get("Manager", "")

    # Direct reports
    for report in contact.get("DirectReports", []):
        person["direct_reports"].append({
            "name": report.get("Name", ""),
            "email": report.get("EmailAddress", ""),
        })

    return person


@mcp.tool()
def find_person(query: str, ctx: Context) -> str:
    """Search for people in the corporate directory.

    Looks up employees by name, email, department, or keyword using the
    Exchange ResolveNames API against Active Directory.

    Args:
        query: Name, email address, or keyword to search for.

    Returns:
        JSON array of matching people with contact details (name, email,
        job_title, department, company, office, phones, address, manager,
        direct_reports, alias).
    """
    client = _get_client(ctx)

    try:
        resolutions = client.resolve_names(query)
    except Exception as e:
        return json.dumps({"error": str(e)})

    if not resolutions:
        return json.dumps([])

    people = [_parse_person(r) for r in resolutions]
    return json.dumps(people, ensure_ascii=False)
