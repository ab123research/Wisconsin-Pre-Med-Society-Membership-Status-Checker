"""Business-rule tests. Run with: python -m unittest -v test_membership."""

from datetime import datetime, timezone
from io import BytesIO
import unittest
from unittest.mock import patch

from membership import (
    Activity, DuplicateMember, Member, MemberNotFound, MembershipBook,
    MembershipError, DEFAULT_WORKSHEET, download_workbook, parse_dues,
    parse_flag, read_workbook,
)

HEADERS = ["Name", "Active\nMember?", "Dues Paid?", "9/23: Fall Kickoff",
           "10/7: Med Student", "10/21: Family Med", "11/4: Psych",
           "11/18: Oncology", "12/2: Game Night", "Total Meetings \nAttended (3)",
           "Volunteered?", "Social Event\nAttended?", "Mentorship\nProgram?",
           "Attended\nOffice Hours?", "Workshop\nAttended?", "Fundraised?", None]


def sample_member(**changes):
    values = dict(name="Example Member", row_number=2, official_active=False,
                  dues_paid=True, dues_record="Year",
                  meetings=tuple(Activity(f"Meeting {i + 1}", i < 2) for i in range(6)),
                  events={key: Activity(label, key in {"social", "office_hours"}) for key, label in (
                      ("volunteering", "Volunteering"), ("social", "Social event"),
                      ("mentorship", "Mentorship program"), ("office_hours", "Office hours"),
                      ("workshop", "Workshop"), ("fundraising", "Fundraising"))}, issues=())
    values.update(changes)
    return Member(**values)


def sample_row(name="Example Member"):
    return [name, False, "Year", "x", "X", "", "", "", "", 99,
            False, True, False, True, False, False, None]


class FakeSheet:
    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, values_only=True):
        return iter(self.rows)


class FakeWorkbook:
    sheetnames = [DEFAULT_WORKSHEET]

    def __init__(self, rows):
        self.sheet = FakeSheet(rows)
        self.closed = False

    def __getitem__(self, name):
        return self.sheet

    def close(self):
        self.closed = True


class MembershipTests(unittest.TestCase):
    def test_attendance_values(self):
        for value in (True, 1, 1.0, "x", " X ", "TRUE", "Yes", "✓"):
            self.assertIs(parse_flag(value), True)
        for value in (False, 0, "false", "No", "", None):
            self.assertIs(parse_flag(value), False)
        for value in (2, -1, "pending", "=A1"):
            self.assertIsNone(parse_flag(value))

    def test_semester_and_year_dues_apply_to_current_term(self):
        for value in ("Year", "Semester", "Yes", True):
            self.assertIs(parse_dues(value), True)
        self.assertIs(parse_dues("No"), False)
        self.assertIsNone(parse_dues(None))
        self.assertIsNone(parse_dues("Pending"))

    def test_three_meeting_boundary_and_denominator(self):
        member = sample_member()
        self.assertEqual(member.meetings_attended, 2)
        self.assertIs(member.meeting_requirement, False)
        self.assertIn("Attend 1 more general meeting.", member.remaining)
        for count in (3, 4, 6):
            member = sample_member(meetings=tuple(Activity(str(i), i < count) for i in range(6)))
            self.assertIs(member.meeting_requirement, True)
            self.assertEqual(member.report()["meetings_total"], 6)

    def test_purple_needs_one_but_keeps_individual_attendance(self):
        member = sample_member()
        self.assertIs(member.purple_met, True)
        self.assertIs(member.events["volunteering"].attended, False)
        self.assertIs(member.events["social"].attended, True)

    def test_each_pink_option_satisfies_group(self):
        for key in ("mentorship", "office_hours", "workshop"):
            events = dict(sample_member().events)
            for option in ("mentorship", "office_hours", "workshop"):
                events[option] = Activity(option, option == key)
            self.assertIs(sample_member(events=events).pink_met, True)

    def test_volunteering_and_social_do_not_replace_fundraising(self):
        events = dict(sample_member().events)
        events["volunteering"] = Activity("Volunteering", True)
        member = sample_member(events=events)
        self.assertIs(member.purple_met, True)
        self.assertIs(member.events["fundraising"].attended, False)
        self.assertIn("Participate in fundraising (orange).", member.remaining)

    def test_official_active_status_is_never_recomputed(self):
        events = {key: Activity(activity.label, True) for key, activity in sample_member().events.items()}
        member = sample_member(events=events, meetings=tuple(Activity(str(i), True) for i in range(6)))
        self.assertEqual(member.requirements_met, 5)
        self.assertEqual(member.remaining, [])
        self.assertIs(member.official_active, False)
        active = sample_member(official_active=True, dues_paid=False)
        self.assertIs(active.official_active, True)
        self.assertTrue(active.remaining)

    def test_unknown_values_are_not_silently_failed(self):
        events = dict(sample_member().events)
        events["volunteering"] = Activity("Volunteering", None)
        events["social"] = Activity("Social event", False)
        member = sample_member(events=events, meetings=(Activity("Meeting 1", None),))
        self.assertIsNone(member.purple_met)
        self.assertIsNone(member.meeting_requirement)
        self.assertTrue(any("secretary" in item for item in member.remaining))

    def test_name_matching_handles_case_whitespace_and_unicode(self):
        book = MembershipBook((sample_member(name="Jamie O’Neil"),), DEFAULT_WORKSHEET, datetime.now(timezone.utc))
        self.assertEqual(book.find("  JAMIE   O'NEIL ").name, "Jamie O’Neil")
        for query in ("Jamie", "", "Someone Else"):
            with self.assertRaises(MemberNotFound):
                book.find(query)

    def test_duplicate_names_are_not_silently_resolved(self):
        book = MembershipBook((sample_member(), sample_member(row_number=3)), DEFAULT_WORKSHEET, datetime.now(timezone.utc))
        with self.assertRaises(DuplicateMember):
            book.find("example member")

    def test_read_uses_meeting_marks_not_stale_total_and_skips_footer(self):
        workbook = FakeWorkbook([HEADERS, sample_row(), sample_row(None), sample_row("totals ")])
        with patch("membership.load_workbook", return_value=workbook):
            book = read_workbook(b"fixture")
        self.assertTrue(workbook.closed)
        self.assertEqual(len(book.members), 1)
        self.assertEqual(book.members[0].meetings_attended, 2)
        self.assertEqual(len(book.members[0].meetings), 6)

    def test_inserted_meeting_updates_denominator(self):
        headers = list(HEADERS)
        headers.insert(9, "12/9: Extra general meeting")
        row = sample_row()
        row.insert(9, "x")
        with patch("membership.load_workbook", return_value=FakeWorkbook([headers, row])):
            member = read_workbook(b"fixture").members[0]
        self.assertEqual(member.meetings_attended, 3)
        self.assertEqual(len(member.meetings), 7)

    def test_missing_headers_and_worksheet_fail_clearly(self):
        with patch("membership.load_workbook", return_value=FakeWorkbook([HEADERS, sample_row()])):
            with self.assertRaisesRegex(MembershipError, "worksheet"):
                read_workbook(b"fixture", "Missing tab")
        headers = list(HEADERS)
        headers[2] = "Wrong heading"
        with patch("membership.load_workbook", return_value=FakeWorkbook([headers, sample_row()])):
            with self.assertRaisesRegex(MembershipError, "Dues Paid"):
                read_workbook(b"fixture")

    def test_invalid_workbook_is_readable_error(self):
        with self.assertRaisesRegex(MembershipError, "valid Excel"):
            read_workbook(b"This is not an Excel file")

    def test_html_download_response_is_rejected(self):
        with patch("membership.urlopen", return_value=BytesIO(b"<html>sign in</html>")):
            with self.assertRaisesRegex(MembershipError, "did not return an Excel"):
                download_workbook()

    def test_sheet_id_cannot_be_arbitrary_url(self):
        with self.assertRaisesRegex(MembershipError, "invalid"):
            download_workbook("https://example.com/")


if __name__ == "__main__":
    unittest.main()
