import os
import sys
import tempfile
import unittest
import uuid
from unittest.mock import patch


PROJECT_ROOT = os.path.dirname(os.path.dirname(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

os.environ.setdefault("SECRET_KEY", "test-secret-key")
os.environ.setdefault(
    "DATABASE_PATH",
    os.path.join(tempfile.mkdtemp(prefix="outlookemail-tests-"), "test.db"),
)

import web_outlook_app as app_module


ACTIVE_ACCOUNT = {
    "id": 1,
    "email": "user@outlook.com",
    "status": "active",
    "account_type": "outlook",
    "provider": "outlook",
    "imap_host": "outlook.live.com",
    "imap_password": "not-used-by-graph",
}

IMAP_ACCOUNT = {
    "id": 3,
    "email": "imap-user@example.com",
    "status": "active",
    "account_type": "imap",
    "provider": "custom",
    "imap_host": "imap.example.com",
    "imap_password": "imap-secret",
    "imap_port": 993,
}

OAUTH_IMAP_ACCOUNT = {
    "id": 4,
    "email": "oauth-imap-user@outlook.com",
    "status": "active",
    "account_type": "imap",
    "provider": "outlook",
    "client_id": "oauth-client-id",
    "refresh_token": "oauth-refresh-token",
    "imap_host": "outlook.live.com",
    "imap_password": "unused-in-oauth-imap",
    "imap_port": 993,
}

GROUP_RECORD = {
    "id": 2,
    "name": "注册组",
    "description": "用于注册流程",
    "color": "#123456",
    "is_system": 0,
    "proxy_url": "",
}

GROUP_ACCOUNTS = [
    {
        "id": 11,
        "email": "group-a@outlook.com",
        "status": "active",
        "group_id": 2,
        "group_name": "注册组",
        "group_color": "#123456",
        "remark": "A",
        "provider": "outlook",
        "pool_status": "available",
        "last_refresh_at": "2026-04-10T10:10:00+00:00",
    },
    {
        "id": 12,
        "email": "group-b@outlook.com",
        "status": "inactive",
        "group_id": 2,
        "group_name": "注册组",
        "group_color": "#123456",
        "remark": "B",
        "provider": "outlook",
        "pool_status": "retired",
        "last_refresh_at": "2026-04-10T10:20:00+00:00",
    },
]


def mock_read_account_messages(account, folder="inbox", **kwargs):
    messages = {
        "inbox": {
            "id": "msg-inbox",
            "date": "2026-04-10T09:00:00+00:00",
            "subject": "Inbox verification code",
        },
        "junkemail": {
            "id": "msg-junk",
            "date": "2026-04-10T10:00:00+00:00",
            "subject": "Junk verification code",
        },
    }
    return {
        "success": True,
        "method": "mock",
        "emails": [messages[folder]],
    }


def mock_read_account_messages_rfc822_utc(account, folder="inbox", **kwargs):
    messages = {
        "inbox": {
            "id": "msg-inbox-rfc822",
            "date": "Wed, 08 Apr 2026 15:01:42 +0000 (UTC)",
            "subject": "Inbox verification code",
        },
        "junkemail": {
            "id": "msg-junk-rfc822",
            "date": "Mon, 13 Apr 2026 16:21:15 +0000 (UTC)",
            "subject": "Junk verification code",
        },
    }
    return {
        "success": True,
        "method": "mock",
        "emails": [messages[folder]],
    }


def mock_read_account_message_detail(account, message_id, folder="inbox"):
    details = {
        "msg-inbox": {
            "id": "msg-inbox",
            "subject": "Inbox verification code",
            "content": "Your verification code is 111111.",
            "html_content": "",
            "from_address": "sender@example.com",
            "timestamp": "2026-04-10T09:00:00+00:00",
            "method": "mock",
        },
        "msg-junk": {
            "id": "msg-junk",
            "subject": "Junk verification code",
            "content": "Your verification code is 654321.",
            "html_content": "",
            "from_address": "sender@example.com",
            "timestamp": "2026-04-10T10:00:00+00:00",
            "method": "mock",
        },
    }
    return {
        "success": True,
        "email": details[message_id],
    }


def mock_read_account_message_detail_rfc822_utc(account, message_id, folder="inbox"):
    details = {
        "msg-inbox-rfc822": {
            "id": "msg-inbox-rfc822",
            "subject": "Inbox verification code",
            "content": "Your verification code is 111111.",
            "html_content": "",
            "from_address": "sender@example.com",
            "timestamp": "Wed, 08 Apr 2026 15:01:42 +0000 (UTC)",
            "method": "mock",
        },
        "msg-junk-rfc822": {
            "id": "msg-junk-rfc822",
            "subject": "Junk verification code",
            "content": "Your verification code is 654321.",
            "html_content": "",
            "from_address": "sender@example.com",
            "timestamp": "Mon, 13 Apr 2026 16:21:15 +0000 (UTC)",
            "method": "mock",
        },
    }
    return {
        "success": True,
        "email": details[message_id],
    }


def mock_read_account_messages_latest_not_code(account, folder="inbox", **kwargs):
    messages = {
        "inbox": [
            {
                "id": "msg-latest-no-code",
                "date": "2026-04-10T10:30:00+00:00",
                "subject": "Security alert",
            },
            {
                "id": "msg-older-with-code",
                "date": "2026-04-10T10:00:00+00:00",
                "subject": "Your verification code",
            },
        ],
        "junkemail": [],
    }
    return {
        "success": True,
        "method": "mock",
        "emails": messages[folder],
    }


def mock_read_account_message_detail_latest_not_code(account, message_id, folder="inbox"):
    details = {
        "msg-latest-no-code": {
            "id": "msg-latest-no-code",
            "subject": "Security alert",
            "content": "There was a login attempt from a new device.",
            "html_content": "",
            "from_address": "security@example.com",
            "timestamp": "2026-04-10T10:30:00+00:00",
            "method": "mock",
        },
        "msg-older-with-code": {
            "id": "msg-older-with-code",
            "subject": "Your verification code",
            "content": "Your verification code is 123456.",
            "html_content": "",
            "from_address": "sender@example.com",
            "timestamp": "2026-04-10T10:00:00+00:00",
            "method": "mock",
        },
    }
    return {
        "success": True,
        "email": details[message_id],
    }


def mock_get_emails_imap_generic(email_addr, imap_password, imap_host, imap_port=993, folder="inbox", provider="custom", skip=0, top=20, proxy_url=""):
    messages = {
        "inbox": {
            "id": "imap-inbox",
            "date": "2026-04-10T09:00:00+00:00",
            "subject": "IMAP inbox verification code",
            "body_preview": "Inbox code 111111",
        },
        "junkemail": {
            "id": "imap-junk",
            "date": "2026-04-10T10:00:00+00:00",
            "subject": "IMAP junk verification code",
            "body_preview": "Junk code 654321",
        },
    }
    return {
        "success": True,
        "emails": [messages[folder]],
        "method": "IMAP (Generic)",
        "has_more": False,
    }


def mock_get_email_detail_imap_generic_result(email_addr, imap_password, imap_host, imap_port=993, message_id="", folder="inbox", provider="custom", proxy_url=""):
    details = {
        "imap-inbox": {
            "id": "imap-inbox",
            "subject": "IMAP inbox verification code",
            "body": "Your verification code is 111111.",
            "body_type": "text",
            "from": "sender@example.com",
            "date": "2026-04-10T09:00:00+00:00",
        },
        "imap-junk": {
            "id": "imap-junk",
            "subject": "IMAP junk verification code",
            "body": "Your verification code is 654321.",
            "body_type": "text",
            "from": "sender@example.com",
            "date": "2026-04-10T10:00:00+00:00",
        },
    }
    return {
        "success": True,
        "email": details[message_id],
    }


def mock_get_emails_imap_with_server(account, client_id, refresh_token, folder="inbox", skip=0, top=20, server="", proxy_url=None, fallback_proxy_urls=None):
    messages = {
        "inbox": {
            "id": "oauth-imap-inbox",
            "date": "2026-04-10T09:00:00+00:00",
            "subject": "OAuth IMAP inbox verification code",
            "body_preview": "Inbox code 777777",
        },
        "junkemail": {
            "id": "oauth-imap-junk",
            "date": "2026-04-10T10:00:00+00:00",
            "subject": "OAuth IMAP junk verification code",
            "body_preview": "Junk code 888888",
        },
    }
    return {
        "success": True,
        "emails": [messages[folder]],
        "method": "IMAP (OAuth)",
        "has_more": False,
    }


def mock_get_email_detail_imap(account, client_id, refresh_token, message_id, folder="inbox", proxy_url=None, fallback_proxy_urls=None):
    details = {
        "oauth-imap-inbox": {
            "id": "oauth-imap-inbox",
            "subject": "OAuth IMAP inbox verification code",
            "from": "sender@example.com",
            "to": "oauth-imap-user@outlook.com",
            "cc": "",
            "date": "2026-04-10T09:00:00+00:00",
            "body": "Your verification code is 777777.",
        },
        "oauth-imap-junk": {
            "id": "oauth-imap-junk",
            "subject": "OAuth IMAP junk verification code",
            "from": "sender@example.com",
            "to": "oauth-imap-user@outlook.com",
            "cc": "",
            "date": "2026-04-10T10:00:00+00:00",
            "body": "Your verification code is 888888.",
        },
    }
    return details[message_id]


class RemovedVerificationApiTests(unittest.TestCase):
    def setUp(self):
        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()

    def test_verification_and_pool_api_routes_are_not_registered(self):
        registered_paths = {rule.rule for rule in app_module.app.url_map.iter_rules()}
        paths = (
            "/api/accounts/1/latest-verification-code",
            "/api/external/verification-code",
            "/api/external/pool/groups",
        )
        for path in paths:
            with self.subTest(path=path):
                self.assertNotIn(path, registered_paths)
class RefreshAccountStatusTests(unittest.TestCase):
    def setUp(self):
        app_module.app.config["TESTING"] = True
        app_module.app.config["WTF_CSRF_ENABLED"] = False
        self.client = app_module.app.test_client()

        with app_module.app.app_context():
            db = app_module.get_db()
            db.execute("DELETE FROM account_refresh_logs")
            db.execute("DELETE FROM accounts")
            db.commit()

    def _insert_outlook_account(self, email_prefix, status="active", account_type="outlook"):
        email = f"{email_prefix}-{uuid.uuid4().hex[:8]}@example.com"
        with app_module.app.app_context():
            db = app_module.get_db()
            encrypted_refresh_token = app_module.encrypt_data("refresh-token-for-test")
            encrypted_password = app_module.encrypt_data("password-for-test")
            db.execute(
                """
                INSERT INTO accounts (
                    email, password, client_id, refresh_token, group_id,
                    remark, status, account_type, provider,
                    imap_host, imap_port, imap_password, forward_enabled
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    email,
                    encrypted_password,
                    "test-client-id",
                    encrypted_refresh_token,
                    1,
                    "",
                    status,
                    account_type,
                    "outlook" if account_type == "outlook" else "custom",
                    "outlook.live.com" if account_type == "outlook" else "imap.example.com",
                    993,
                    "",
                    0,
                ),
            )
            db.commit()
            row = db.execute("SELECT id, email FROM accounts WHERE email = ?", (email,)).fetchone()
            return row["id"], row["email"]

    def test_log_refresh_result_disables_failed_and_reactivates_success(self):
        account_id, email = self._insert_outlook_account("refresh-status")

        with app_module.app.app_context():
            app_module.log_refresh_result(account_id, email, "manual", "failed", "token invalid")
            db = app_module.get_db()
            row = db.execute("SELECT status FROM accounts WHERE id = ?", (account_id,)).fetchone()
            self.assertEqual(row["status"], "inactive")

            app_module.log_refresh_result(account_id, email, "manual", "success", None)
            row = db.execute("SELECT status, last_refresh_at FROM accounts WHERE id = ?", (account_id,)).fetchone()
            self.assertEqual(row["status"], "active")
            self.assertIsNotNone(row["last_refresh_at"])

    def test_load_refresh_target_accounts_includes_failed_inactive_accounts(self):
        active_id, _ = self._insert_outlook_account("active-outlook", status="active", account_type="outlook")
        inactive_failed_id, inactive_failed_email = self._insert_outlook_account("inactive-failed", status="inactive", account_type="outlook")
        inactive_nonfailed_id, inactive_nonfailed_email = self._insert_outlook_account("inactive-ok", status="inactive", account_type="outlook")
        imap_active_id, _ = self._insert_outlook_account("imap-active", status="active", account_type="imap")

        with app_module.app.app_context():
            db = app_module.get_db()
            db.execute(
                "INSERT INTO account_refresh_logs (account_id, account_email, refresh_type, status, error_message) VALUES (?, ?, ?, ?, ?)",
                (inactive_failed_id, inactive_failed_email, "manual", "failed", "x"),
            )
            db.execute(
                "INSERT INTO account_refresh_logs (account_id, account_email, refresh_type, status, error_message) VALUES (?, ?, ?, ?, ?)",
                (inactive_nonfailed_id, inactive_nonfailed_email, "manual", "success", None),
            )
            db.commit()

            targets = app_module.load_refresh_target_accounts(db)
            target_ids = {row["id"] for row in targets}

            self.assertIn(active_id, target_ids)
            self.assertIn(inactive_failed_id, target_ids)
            self.assertNotIn(inactive_nonfailed_id, target_ids)
            self.assertNotIn(imap_active_id, target_ids)

    def test_refresh_failed_api_reactivates_inactive_account_on_success(self):
        account_id, email = self._insert_outlook_account("retry-inactive", status="inactive", account_type="outlook")

        with app_module.app.app_context():
            db = app_module.get_db()
            db.execute(
                "INSERT INTO account_refresh_logs (account_id, account_email, refresh_type, status, error_message) VALUES (?, ?, ?, ?, ?)",
                (account_id, email, "manual", "failed", "token expired"),
            )
            db.commit()

        with patch.object(app_module, "probe_account_permissions", return_value={
            "success": True,
            "health_status": "graph_only",
            "graph_status": "ok",
            "imap_status": "fail",
            "permission_scopes": {"graph": "Mail.Read", "imap": ""},
            "error": "IMAP unavailable",
        }):
            with self.client.session_transaction() as session:
                session["logged_in"] = True

            response = self.client.post("/api/accounts/refresh-failed")
            payload = response.get_json()

        self.assertEqual(response.status_code, 200)
        self.assertTrue(payload["success"])
        self.assertEqual(payload["success_count"], 1)
        self.assertEqual(payload["failed_count"], 0)

        with app_module.app.app_context():
            db = app_module.get_db()
            row = db.execute("SELECT status FROM accounts WHERE id = ?", (account_id,)).fetchone()
            self.assertEqual(row["status"], "active")


if __name__ == "__main__":
    unittest.main()
