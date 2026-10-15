"""Streamlit interface checks with synthetic member records."""

from datetime import datetime, timezone
from pathlib import Path
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from membership import Activity, DEFAULT_WORKSHEET, MembershipBook
from test_membership import sample_member

APP = str(Path(__file__).with_name("app.py"))


def html_values(app):
    return "\n".join(element.proto.body for element in app.get("html"))


class AppTests(unittest.TestCase):
    def setUp(self):
        from streamlit import cache_data
        cache_data.clear()

    def test_form_and_three_column_result(self):
        book = MembershipBook((sample_member(),), DEFAULT_WORKSHEET, datetime.now(timezone.utc))
        with patch("membership.load_membership", return_value=book):
            app = AppTest.from_file(APP).run(timeout=10)
            self.assertEqual(len(app.exception), 0)
            app.text_input[0].set_value("  EXAMPLE   MEMBER ")
            app.button[0].click().run(timeout=10)
            self.assertEqual(len(app.exception), 0)
            html = html_values(app)
            self.assertIn("Inactive member", html)
            self.assertIn("2 / 6", html)
            self.assertIn("Volunteering", html)
            self.assertIn("Social event", html)
            self.assertIn("<th scope=\"col\">Met</th>", html)
            self.assertIn("<th scope=\"col\">Requirement</th>", html)
            self.assertIn("<th scope=\"col\">Attendance</th>", html)
            self.assertIn("Attend 1 more general meeting.", html)
            self.assertIn("Participate in fundraising (orange).", html)

    def test_unknown_and_blank_names_show_clear_messages(self):
        book = MembershipBook((sample_member(),), DEFAULT_WORKSHEET, datetime.now(timezone.utc))
        with patch("membership.load_membership", return_value=book):
            app = AppTest.from_file(APP).run()
            app.button[0].click().run()
            self.assertIn("Enter your full name", app.warning[0].value)
            app.text_input[0].set_value("Missing Person")
            app.button[0].click().run()
            self.assertIn("No matching member", app.error[0].value)
            self.assertNotIn("Example Member", html_values(app))

    def test_html_in_member_name_is_escaped(self):
        name = '<img src=x onerror="alert(1)">'
        book = MembershipBook((sample_member(name=name),), DEFAULT_WORKSHEET, datetime.now(timezone.utc))
        with patch("membership.load_membership", return_value=book):
            app = AppTest.from_file(APP).run()
            app.text_input[0].set_value(name)
            app.button[0].click().run()
            self.assertEqual(len(app.exception), 0)
            html = html_values(app)
            self.assertNotIn('<img src=x', html)
            self.assertIn('&lt;img src=x', html)

    def test_requirements_do_not_override_official_status(self):
        member = sample_member(meetings=tuple(Activity(str(i), True) for i in range(6)))
        member = sample_member(meetings=member.meetings,
                               events={key: Activity(activity.label, True) for key, activity in member.events.items()})
        book = MembershipBook((member,), DEFAULT_WORKSHEET, datetime.now(timezone.utc))
        with patch("membership.load_membership", return_value=book):
            app = AppTest.from_file(APP).run()
            app.text_input[0].set_value(member.name)
            app.button[0].click().run()
            self.assertEqual(len(app.exception), 0)
            html = html_values(app)
            self.assertIn("Inactive member", html)
            self.assertIn("5 of 5 requirements met", html)
            self.assertIn("Ask the secretary to review", html)


if __name__ == "__main__":
    unittest.main()
