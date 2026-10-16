"""Streamlit entrypoint: streamlit run app.py."""

from __future__ import annotations

from html import escape
import os
from zoneinfo import ZoneInfo

import streamlit as st

from branding import brand_css, hero_html, form_heading_html, footer_html

from membership import (
    DEFAULT_WORKSHEET, REQUIRED_MEETINGS, ConnectionSettings,
    MembershipError, Member, load_membership, state_text, parse_connection_settings,
)

TERM = os.environ.get("MEMBERSHIP_TERM", "Fall 2026")
WORKSHEET = os.environ.get("WORKSHEET_NAME", DEFAULT_WORKSHEET)
LOCAL_FILE = os.environ.get("MEMBERSHIP_FILE") or None

CSS = brand_css()


@st.cache_data(ttl=60, max_entries=4, show_spinner=False)
def get_book(settings: ConnectionSettings | None, local_file: str | None):
    # Credentials and the full roster stay on the server.
    if local_file:
        return load_membership(file=local_file, worksheet=WORKSHEET)
    if settings is None:
        raise MembershipError("Private-sheet settings are missing.")
    return load_membership(sheet_id=settings.spreadsheet_id, worksheet=settings.worksheet,
                           service_account_info=settings.service_account_info)


def private_connection_settings() -> ConnectionSettings:
    try:
        section = dict(st.secrets["google_sheets"])
    except Exception:
        # Streamlit/TOML parsing errors can include secret text; redact at this boundary.
        raise MembershipError("Private-sheet settings could not be read. Check the google_sheets section in .streamlit/secrets.toml locally, or in the app's Streamlit Secrets settings when deployed.") from None
    return parse_connection_settings(section)


def marker(state: bool | None) -> str:
    symbol, css, label = ("✓", "yes", "Met / attended") if state is True else (
        ("✗", "no", "Not met / not attended") if state is False else ("?", "unknown", "Needs review"))
    return f'<span class="mark {css}" role="img" aria-label="{label}">{symbol}</span>'


def table_row(state: bool | None, label: str, attendance: str, *,
              group: str = "", note: str = "") -> str:
    css = f"group {group}" if group else "detail"
    extra = f'<span class="note">{escape(note)}</span>' if note else ""
    return (f'<tr class="{css}"><td>{marker(state)}</td>'
            f'<td>{escape(label)}{extra}</td><td>{escape(attendance)}</td></tr>')


def requirements_html(member: Member) -> str:
    attendance = lambda flag: state_text(flag, "Attended", "Not Attended")
    rows = [table_row(member.dues_paid, "Semester dues", state_text(member.dues_paid, "Paid", "Not paid"),
                      group="dues", note="Dues payment recorded for this semester.")]
    rows.append(table_row(member.meeting_requirement, "General meetings · Blue",
                          f"{member.meetings_attended} / {len(member.meetings)}",
                          group="blue", note=f"Attend at least {REQUIRED_MEETINGS} general meetings."))
    rows.append(table_row(member.purple_met, "Volunteer or social event · Purple",
                          attendance(member.purple_met), group="purple", note="Attend at least one of the two activities below."))
    for key in ("volunteering", "social"):
        activity = member.events[key]
        rows.append(table_row(activity.attended, activity.label, attendance(activity.attended)))
    rows.append(table_row(member.pink_met, "Mentorship, office hours or workshop · Pink",
                          attendance(member.pink_met), group="pink", note="Complete at least one of the three activities below."))
    for key in ("mentorship", "office_hours", "workshop"):
        activity = member.events[key]
        rows.append(table_row(activity.attended, activity.label, attendance(activity.attended)))
    fund = member.events["fundraising"]
    rows.append(table_row(fund.attended, "Fundraising · Orange", attendance(fund.attended),
                          group="orange", note="Participate in at least one fundraiser."))
    return ('<div class="membership-table-wrap"><table class="membership-table" aria-label="Membership requirements">'
            '<thead><tr><th scope="col">Met</th><th scope="col">Requirement</th>'
            '<th scope="col">Attendance</th></tr></thead><tbody>' + "".join(rows) + '</tbody></table></div>')


def profile_html(member: Member) -> str:
    official = state_text(member.official_active, "✓ Active member", "✗ Inactive member")
    state_class = "is-active" if member.official_active is True else (
        "is-unknown" if member.official_active is None else "is-inactive")
    circumference = 251.3274
    completed_arc = circumference * member.requirements_met / 5
    ring = ('<svg class="wps-profile-ring" viewBox="0 0 90 90" aria-hidden="true">'
            '<circle cx="45" cy="45" r="40" fill="none" stroke="#F0E9E4" stroke-width="5"/>'
            f'<circle cx="45" cy="45" r="40" fill="none" stroke="#C5050C" stroke-width="5" '
            f'stroke-dasharray="{completed_arc:.3f} {circumference - completed_arc:.3f}" '
            'transform="rotate(-90 45 45)"/></svg>')
    return (f'<section class="membership-profile"><div><div class="label">Official membership status</div>'
            f'<h2>{escape(member.name)}</h2><div class="official {state_class}">{escape(official)}</div>'
            '<p>As recorded by the secretary in the membership sheet.</p></div>'
            '<div class="wps-profile-emblem" '
            f'role="img" aria-label="{member.requirements_met} of 5 requirements met">'
            f'{ring}'
            f'<span>{member.requirements_met}<span aria-hidden="true">/5</span><small>Complete</small></span>'
            '</div></section>')


def show_member(member: Member) -> None:
    st.html(profile_html(member))
    st.html(f'<p class="membership-count">{member.requirements_met} of 5 requirements met'
            '<span>Your semester at a glance</span></p>' + requirements_html(member))
    if member.remaining:
        items = "".join(f"<li>{escape(item)}</li>" for item in member.remaining)
        st.html('<section class="membership-next"><h3>What remains</h3><ul>' + items + '</ul></section>')
    else:
        text = "You have met all listed requirements."
        if member.official_active is not True:
            text += " Ask the secretary to review your official membership status."
        st.html(f'<section class="membership-next"><h3>All listed requirements met</h3><p>{escape(text)}</p></section>')
    if member.official_active is True and member.remaining:
        st.caption("Your official status is Active. The breakdown above shows the current entries; contact the secretary about any differences.")
    if member.issues:
        st.warning("Some entries need the secretary’s review. A question mark means the value could not be interpreted.")
        with st.expander("Entries to review"):
            for issue in member.issues:
                st.text(issue)
    with st.expander("General meeting attendance"):
        rows = "".join(table_row(activity.attended, activity.label,
                               state_text(activity.attended, "Attended", "Not Attended"))
                       for activity in member.meetings)
        st.html('<div class="membership-table-wrap"><table class="membership-table" aria-label="General meeting attendance">'
                '<thead><tr><th scope="col">Met</th><th scope="col">Meeting</th>'
                '<th scope="col">Attendance</th></tr></thead><tbody>' + rows + '</tbody></table></div>')
    st.caption("Volunteering and social attendance are listed separately. Exceptions and corrections are handled by the secretary.")


def main() -> None:
    st.set_page_config(page_title=f"WPS membership | {TERM}", page_icon="🩺", layout="centered")
    st.html(CSS)
    st.html(hero_html(TERM))
    try:
        show_lookup()
    finally:
        st.html(footer_html(TERM))


def show_lookup() -> None:
    with st.form("member_lookup"):
        st.html(form_heading_html())
        name = st.text_input("Your full name", placeholder="Enter your name as listed on the sheet", max_chars=150)
        submitted = st.form_submit_button("Check my membership", type="primary", use_container_width=True)
    if submitted:
        st.session_state["member_query"] = name.strip()
    if "member_query" not in st.session_state:
        st.caption("Use your full name. Capitalization and extra spaces do not matter.")
        return
    query = st.session_state["member_query"]
    if not query:
        st.warning("Enter your full name to check your membership.")
        return
    if st.button("Refresh from sheet", help="Get the latest entries from the membership sheet."):
        get_book.clear()
    try:
        with st.spinner("Checking the membership sheet…"):
            settings = None if LOCAL_FILE else private_connection_settings()
            book = get_book(settings, LOCAL_FILE)
            member = book.find(query)
    except MembershipError as exc:
        st.error(str(exc))
        return
    show_member(member)
    fetched = book.fetched_at.astimezone(ZoneInfo("America/Chicago"))
    source = "Local workbook" if LOCAL_FILE else "Live membership sheet"
    st.caption(f"{source} checked {fetched.strftime('%b %d, %Y at %I:%M %p %Z')}. "
               "Sheet data refreshes on the next check after 60 seconds, or when you select Refresh from sheet.")


if __name__ == "__main__":
    main()
