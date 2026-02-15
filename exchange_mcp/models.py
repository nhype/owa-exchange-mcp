"""Return type definitions for MCP tools.

Uses TypedDict for lightweight structured return types without
requiring Pydantic at the model layer.
"""

from typing import TypedDict


class Email(TypedDict, total=False):
    """An email message."""
    subject: str
    sender: str
    sender_name: str
    date: str
    is_read: bool
    has_attachments: bool
    has_links: bool
    item_id: str
    size: int
    is_meeting: bool
    item_type: str
    to: list[str]
    cc: list[str]
    body: str
    body_type: str
    attachments: list[dict]
    # Meeting-specific fields
    location: str
    start: str
    end: str
    attendees_required: list[str]
    attendees_optional: list[str]


class CalendarEvent(TypedDict, total=False):
    """A calendar event / meeting."""
    subject: str
    start: str
    end: str
    location: str
    is_all_day: bool
    is_cancelled: bool
    is_meeting: bool
    is_recurring: bool
    organizer: str
    organizer_email: str
    my_response: str
    item_id: str
    body: str
    attendees_required: list[str]
    attendees_optional: list[str]


class Person(TypedDict, total=False):
    """A person from the Exchange directory."""
    name: str
    email: str
    mailbox_type: str
    first_name: str
    last_name: str
    job_title: str
    department: str
    company: str
    office: str
    alias: str
    manager: str
    manager_email: str
    phones: dict[str, str]
    address: str
    direct_reports: list[dict[str, str]]


class Folder(TypedDict, total=False):
    """A mail folder."""
    name: str
    id: str
    total_count: int
    unread_count: int
    child_folder_count: int


class FreeSlot(TypedDict):
    """A free time slot."""
    date: str
    start: str
    end: str
    duration_minutes: int


class Availability(TypedDict, total=False):
    """Availability data for a single attendee."""
    email: str
    busy_slots: int
    free_slots: int
    merged_freebusy: str


class MeetingResult(TypedDict, total=False):
    """Result of creating/updating a meeting."""
    success: bool
    subject: str
    date: str
    start_time: str
    end_time: str
    location: str
    required_attendees: list[str]
    optional_attendees: list[str]
    error: str
