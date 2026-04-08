"""
Unit tests for EmailClassifier.build_engaged_domains().

Run with:
    python -m pytest test_classifier.py -v
or:
    python test_classifier.py
"""

import unittest
from datetime import datetime, timedelta, timezone

from cleanser import EmailClassifier, load_config


def _make_msg(
    sender_email: str,
    is_read: bool,
    days_ago: int,
    has_unsub_header: bool = False,
    body_preview: str = "",
) -> dict:
    """Build a minimal raw Graph API message dict for testing."""
    received = (datetime.now(timezone.utc) - timedelta(days=days_ago)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    headers = (
        [{"name": "List-Unsubscribe", "value": "<mailto:unsub@example.com>"}]
        if has_unsub_header
        else []
    )
    return {
        "id": f"msg-{sender_email}-{days_ago}",
        "subject": "Test",
        "sender": {"emailAddress": {"name": "Test", "address": sender_email}},
        "receivedDateTime": received,
        "isRead": is_read,
        "conversationId": "conv-1",
        "internetMessageHeaders": headers,
        "bodyPreview": body_preview,
    }


CONFIG = {
    "engagement_window_days": 60,
    "unsubscribe_signals": ["unsubscribe", "opt-out"],
}


class TestBuildEngagedDomains(unittest.TestCase):

    def test_read_with_unsub_header_within_window(self):
        """A read email with List-Unsubscribe header within 60 days → domain included."""
        msgs = [_make_msg("news@example.com", is_read=True, days_ago=10, has_unsub_header=True)]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertIn("example.com", result)

    def test_unread_with_unsub_header_excluded(self):
        """An unread email is not evidence of engagement — domain not included."""
        msgs = [_make_msg("news@example.com", is_read=False, days_ago=10, has_unsub_header=True)]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertNotIn("example.com", result)

    def test_read_outside_window_excluded(self):
        """A read email older than the engagement window → domain not included."""
        msgs = [_make_msg("news@example.com", is_read=True, days_ago=90, has_unsub_header=True)]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertNotIn("example.com", result)

    def test_read_no_unsub_signal_excluded(self):
        """A read email with no unsubscribe signal → not counted as newsletter engagement."""
        msgs = [_make_msg("boss@work.com", is_read=True, days_ago=5)]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertNotIn("work.com", result)

    def test_unsub_signal_in_body_preview(self):
        """Unsubscribe keyword in bodyPreview counts as a signal (no header needed)."""
        msgs = [_make_msg("promo@shop.com", is_read=True, days_ago=5, body_preview="Click here to unsubscribe")]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertIn("shop.com", result)

    def test_multiple_domains(self):
        """Multiple engaged domains are all collected."""
        msgs = [
            _make_msg("a@alpha.com", is_read=True, days_ago=5, has_unsub_header=True),
            _make_msg("b@beta.com", is_read=True, days_ago=5, has_unsub_header=True),
            _make_msg("c@gamma.com", is_read=False, days_ago=5, has_unsub_header=True),  # unread — excluded
        ]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertIn("alpha.com", result)
        self.assertIn("beta.com", result)
        self.assertNotIn("gamma.com", result)

    def test_empty_messages(self):
        """Empty message list returns empty set."""
        result = EmailClassifier.build_engaged_domains([], CONFIG)
        self.assertEqual(result, set())

    def test_exactly_at_window_boundary(self):
        """A message received exactly at the boundary (60 days ago) is excluded (strictly within)."""
        msgs = [_make_msg("news@edge.com", is_read=True, days_ago=60, has_unsub_header=True)]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertNotIn("edge.com", result)

    def test_one_day_inside_window(self):
        """A message received 59 days ago is within the window."""
        msgs = [_make_msg("news@edge.com", is_read=True, days_ago=59, has_unsub_header=True)]
        result = EmailClassifier.build_engaged_domains(msgs, CONFIG)
        self.assertIn("edge.com", result)

    def test_custom_window(self):
        """engagement_window_days config value is respected."""
        cfg = {**CONFIG, "engagement_window_days": 7}
        msgs = [
            _make_msg("news@recent.com", is_read=True, days_ago=5, has_unsub_header=True),
            _make_msg("news@old.com", is_read=True, days_ago=10, has_unsub_header=True),
        ]
        result = EmailClassifier.build_engaged_domains(msgs, cfg)
        self.assertIn("recent.com", result)
        self.assertNotIn("old.com", result)


if __name__ == "__main__":
    unittest.main(verbosity=2)
