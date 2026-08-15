import importlib
import os
import sqlite3
import tempfile
import unittest
from unittest.mock import patch


os.environ.setdefault('SECRET_KEY', 'test-secret-key')
_temp_dir = tempfile.mkdtemp(prefix='outlookEmail-access-profile-tests-')
os.environ['DATABASE_PATH'] = os.path.join(_temp_dir, 'test.db')

app_module = importlib.import_module('web_outlook_app')


class FakeResponse:
    status_code = 200

    @staticmethod
    def json():
        return {'access_token': 'access-token'}


class AccountAccessProfileTests(unittest.TestCase):
    def setUp(self):
        app_module.init_db()

    def test_database_migration_adds_access_profile_columns(self):
        with sqlite3.connect(app_module.DATABASE) as conn:
            columns = {row[1] for row in conn.execute('PRAGMA table_info(accounts)')}
        self.assertIn('mail_protocol', columns)
        self.assertIn('oauth_authority', columns)
        self.assertIn('health_status', columns)
        self.assertIn('graph_status', columns)
        self.assertIn('imap_status', columns)
        self.assertIn('permission_scopes', columns)

    def test_graph_token_uses_consumers_authority(self):
        with patch.object(app_module, 'post_with_proxy_fallback', return_value=FakeResponse()) as request:
            result = app_module.get_access_token_graph_result(
                'client-id', 'refresh-token', oauth_authority='consumers'
            )
        self.assertTrue(result['success'])
        self.assertEqual(
            request.call_args.args[0],
            'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
        )

    def test_imap_token_uses_configured_common_authority(self):
        with patch.object(app_module, 'post_with_proxy_fallback', return_value=FakeResponse()) as request:
            result = app_module.get_access_token_imap_result(
                'client-id', 'refresh-token', oauth_authority='common'
            )
        self.assertTrue(result['success'])
        self.assertEqual(
            request.call_args.args[0],
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        )

    def test_graph_only_does_not_fallback_to_imap(self):
        account = {
            'email': 'graph@example.com',
            'client_id': 'client-id',
            'refresh_token': 'refresh-token',
            'account_type': 'outlook',
            'mail_protocol': 'graph',
            'oauth_authority': 'consumers',
        }
        with patch.object(app_module, 'get_emails_graph', return_value={'success': False, 'error': 'failed'}) as graph, \
                patch.object(app_module, 'get_emails_imap_with_server') as imap:
            result = app_module.fetch_account_folder_emails(account, 'inbox', 0, 20)
        self.assertFalse(result['success'])
        self.assertEqual(graph.call_args.args[-1], 'consumers')
        imap.assert_not_called()

    def test_imap_oauth_skips_graph(self):
        account = {
            'email': 'imap@example.com',
            'client_id': 'client-id',
            'refresh_token': 'refresh-token',
            'account_type': 'outlook',
            'mail_protocol': 'imap_oauth',
            'oauth_authority': 'consumers',
        }
        with patch.object(app_module, 'get_emails_graph') as graph, \
                patch.object(app_module, 'get_emails_imap_with_server', return_value={
                    'success': True, 'emails': [], 'has_more': False
                }) as imap:
            result = app_module.fetch_account_folder_emails(account, 'inbox', 0, 20)
        self.assertTrue(result['success'])
        graph.assert_not_called()
        self.assertEqual(imap.call_args.args[-1], 'consumers')

    def test_auto_falls_back_from_graph_to_imap(self):
        account = {
            'email': 'auto@example.com',
            'client_id': 'client-id',
            'refresh_token': 'refresh-token',
            'account_type': 'outlook',
            'mail_protocol': 'auto',
            'oauth_authority': 'common',
        }
        with patch.object(app_module, 'get_emails_graph', return_value={'success': False, 'error': 'failed'}), \
                patch.object(app_module, 'get_emails_imap_with_server', return_value={
                    'success': True, 'emails': [], 'has_more': False
                }) as imap:
            result = app_module.fetch_account_folder_emails(account, 'inbox', 0, 20)
        self.assertTrue(result['success'])
        self.assertEqual(result['method'], 'IMAP (New)')
        self.assertEqual(imap.call_args.args[-1], 'common')

    def test_permission_probe_classifies_full_access(self):
        class FakeImap:
            def authenticate(self, *_args):
                return 'OK', []

            def select(self, *_args, **_kwargs):
                return 'OK', [b'1']

            def logout(self):
                return 'BYE', []

        token = {'success': True, 'access_token': 'token', 'scope': 'Mail.Read IMAP.AccessAsUser.All'}
        with patch.object(app_module, 'get_access_token_graph_result', return_value=token), \
                patch.object(app_module, 'get_access_token_imap_result', return_value=token), \
                patch.object(app_module, 'get_with_proxy_fallback', return_value=FakeResponse()), \
                patch.object(app_module.imaplib, 'IMAP4_SSL', return_value=FakeImap()):
            result = app_module.probe_account_permissions(
                'user@outlook.com', 'client-id', 'refresh-token',
                mail_protocol='auto', oauth_authority='consumers'
            )
        self.assertTrue(result['success'])
        self.assertEqual(result['health_status'], 'all')
        self.assertEqual(result['graph_status'], 'ok')
        self.assertEqual(result['imap_status'], 'ok')

    def test_permission_probe_respects_configured_protocol(self):
        graph_token = {'success': True, 'access_token': 'token', 'scope': 'Mail.Read'}
        with patch.object(app_module, 'get_access_token_graph_result', return_value=graph_token), \
                patch.object(app_module, 'get_access_token_imap_result', return_value={
                    'success': False, 'error': 'scope missing'
                }), patch.object(app_module, 'get_with_proxy_fallback', return_value=FakeResponse()):
            result = app_module.probe_account_permissions(
                'user@outlook.com', 'client-id', 'refresh-token',
                mail_protocol='imap_oauth', oauth_authority='consumers'
            )
        self.assertFalse(result['success'])
        self.assertEqual(result['health_status'], 'graph_only')


if __name__ == '__main__':
    unittest.main()
