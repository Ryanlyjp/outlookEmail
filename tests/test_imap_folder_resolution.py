import importlib
import json
import os
import sqlite3
import tempfile
import threading
import unittest
from unittest.mock import MagicMock, call, patch


os.environ.setdefault('SECRET_KEY', 'test-secret-key')
_temp_dir = tempfile.mkdtemp(prefix='outlookEmail-tests-')
os.environ['DATABASE_PATH'] = os.path.join(_temp_dir, 'test.db')

web_outlook_app = importlib.import_module('web_outlook_app')


class FakeMail:
    def __init__(self, selectable=None, list_entries=None, selectable_by_mode=None):
        self.selectable = set(selectable or [])
        self.selectable_by_mode = dict(selectable_by_mode or {})
        self.list_entries = list_entries or []
        self.select_calls = []
        self.logged_out = False
        self.xatom_calls = []

    def login(self, *_args, **_kwargs):
        return 'OK', [b'logged in']

    def xatom(self, name, *args):
        self.xatom_calls.append((name, args))
        return 'OK', [b'ID completed']

    def select(self, name, readonly=True):
        self.select_calls.append((name, readonly))
        if (name, readonly) in self.selectable_by_mode:
            return self.selectable_by_mode[(name, readonly)]
        if name in self.selectable:
            return 'OK', [b'']
        return 'NO', [b'folder not found']

    def list(self):
        return 'OK', self.list_entries

    def uid(self, *_args, **_kwargs):
        return 'OK', [b'']

    def logout(self):
        self.logged_out = True
        return 'BYE', [b'logout']


class ImapFolderResolutionTests(unittest.TestCase):
    def test_parse_outlook_import_default_order(self):
        parsed = web_outlook_app.parse_outlook_account_string(
            'user@outlook.com----password123----24d9a0ed-8787-4584-883c-2fd79308940a----0.AXEA_refresh',
            'client_id_refresh_token',
        )

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed['client_id'], '24d9a0ed-8787-4584-883c-2fd79308940a')
        self.assertEqual(parsed['refresh_token'], '0.AXEA_refresh')

    def test_parse_outlook_import_reversed_order_even_when_selector_is_default(self):
        parsed = web_outlook_app.parse_outlook_account_string(
            'user@outlook.com----password123----0.AXEA_refresh----24d9a0ed-8787-4584-883c-2fd79308940a',
            'client_id_refresh_token',
        )

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed['client_id'], '24d9a0ed-8787-4584-883c-2fd79308940a')
        self.assertEqual(parsed['refresh_token'], '0.AXEA_refresh')

    def test_resolve_126_inbox_from_listed_folder(self):
        mail = FakeMail(
            selectable={'INBOX.收件箱'},
            list_entries=[
                b'(\\HasNoChildren) "." "INBOX.Archive"',
                '(\\HasNoChildren) "." "INBOX.收件箱"'.encode('utf-8'),
            ],
        )

        selected, diagnostics = web_outlook_app.resolve_imap_folder(mail, '126', 'inbox', readonly=True)

        self.assertEqual(selected, 'INBOX.收件箱')
        self.assertIn('INBOX.收件箱', diagnostics.get('matched_folders', []))
        self.assertNotIn('INBOX.Archive', diagnostics.get('matched_folders', []))

    def test_resolve_junk_folder_from_terminal_alias(self):
        mail = FakeMail(
            selectable={'INBOX.Spam'},
            list_entries=[b'(\\HasNoChildren) "." "INBOX.Spam"'],
        )

        selected, diagnostics = web_outlook_app.resolve_imap_folder(mail, 'custom', 'junkemail', readonly=True)

        self.assertEqual(selected, 'INBOX.Spam')
        self.assertEqual(diagnostics.get('matched_folders'), ['INBOX.Spam'])

    def test_imap_folder_not_found_returns_available_folders(self):
        mail = FakeMail(
            selectable=set(),
            list_entries=[b'(\\HasNoChildren) "." "INBOX"', b'(\\HasNoChildren) "." "INBOX.Archive"'],
        )

        with patch.object(web_outlook_app, 'create_imap_connection', return_value=mail):
            result = web_outlook_app.get_emails_imap_generic(
                email_addr='user@example.com',
                imap_password='secret',
                imap_host='imap.example.com',
                folder='deleteditems',
                provider='custom',
            )

        self.assertFalse(result['success'])
        self.assertEqual(result['error_code'], 'IMAP_FOLDER_NOT_FOUND')
        details = json.loads(result['error']['details'])
        self.assertEqual(details['folder'], 'deleteditems')
        self.assertEqual(details['provider'], 'custom')
        self.assertEqual(details['available_folders'], ['INBOX', 'INBOX.Archive'])
        self.assertTrue(mail.logged_out)

    def test_fallback_from_examine_to_select_for_126_inbox(self):
        mail = FakeMail(
            selectable_by_mode={
                ('INBOX', True): ('NO', [b'EXAMINE not supported']),
                ('"INBOX"', True): ('NO', [b'EXAMINE not supported']),
                ('INBOX', False): ('OK', [b'12']),
            },
            list_entries=[b'(\\HasNoChildren) "." "INBOX"'],
        )

        selected, diagnostics = web_outlook_app.resolve_imap_folder(mail, '126', 'inbox', readonly=True)

        self.assertEqual(selected, 'INBOX')
        self.assertEqual(diagnostics.get('fallback_mode'), 'select')
        self.assertIn(('INBOX', True), mail.select_calls)
        self.assertIn(('INBOX', False), mail.select_calls)

    def test_unsafe_login_is_classified_as_provider_block(self):
        unsafe = "Unsafe Login. Please contact kefu@188.com for help"
        mail = FakeMail(
            selectable_by_mode={
                ('INBOX', True): ('NO', [unsafe.encode('utf-8')]),
                ('"INBOX"', True): ('NO', [unsafe.encode('utf-8')]),
                ('INBOX', False): ('NO', [unsafe.encode('utf-8')]),
                ('"INBOX"', False): ('NO', [unsafe.encode('utf-8')]),
            },
            list_entries=[b'(\\HasNoChildren) "." "INBOX"'],
        )

        with patch.object(web_outlook_app, 'create_imap_connection', return_value=mail):
            result = web_outlook_app.get_emails_imap_generic(
                email_addr='user@126.com',
                imap_password='secret',
                imap_host='imap.126.com',
                folder='inbox',
                provider='126',
            )

        self.assertFalse(result['success'])
        self.assertEqual(result['error_code'], 'IMAP_UNSAFE_LOGIN_BLOCKED')
        self.assertEqual(result['error']['status'], 403)
        self.assertEqual(result['error']['code'], 'IMAP_UNSAFE_LOGIN_BLOCKED')
        self.assertIn('Unsafe Login', result['error']['message'])
        self.assertEqual(mail.xatom_calls[0][0], 'ID')

    def test_send_imap_id_after_login(self):
        mail = FakeMail(
            selectable={'INBOX'},
            list_entries=[b'(\\HasNoChildren) "." "INBOX"'],
        )

        with patch.object(web_outlook_app, 'create_imap_connection', return_value=mail):
            result = web_outlook_app.get_emails_imap_generic(
                email_addr='user@126.com',
                imap_password='secret',
                imap_host='imap.126.com',
                folder='inbox',
                provider='126',
            )

        self.assertTrue(result['success'])
        self.assertEqual(mail.xatom_calls[0][0], 'ID')
        payload = mail.xatom_calls[0][1][0]
        self.assertIn('"name" "outlookEmail"', payload)
        self.assertIn('"version"', payload)

    def test_fetch_account_emails_all_fetches_in_parallel(self):
        account = {
            'email': 'user@outlook.com',
            'account_type': 'oauth',
        }
        barrier = threading.Barrier(2, timeout=1)

        def fake_fetch(target_account, folder, skip, top, proxy_url='', fallback_proxy_urls=None):
            self.assertIs(target_account, account)
            self.assertEqual(skip, 0)
            self.assertEqual(top, 20)
            self.assertEqual(proxy_url, 'socks5://primary')
            self.assertEqual(fallback_proxy_urls, ['socks5://fallback'])
            try:
                barrier.wait()
            except threading.BrokenBarrierError as exc:
                raise AssertionError('folder=all 仍然是串行抓取') from exc
            return {
                'success': True,
                'emails': [{
                    'id': folder,
                    'folder': folder,
                    'date': '2026-01-01T00:00:00Z' if folder == 'inbox' else '2026-01-02T00:00:00Z',
                }],
                'method': f'method-{folder}',
                'has_more': False,
            }

        with patch.object(web_outlook_app, 'get_account_proxy_url', return_value='socks5://primary'):
            with patch.object(web_outlook_app, 'get_account_proxy_failover_urls', return_value=['socks5://fallback']):
                with patch.object(web_outlook_app, 'fetch_account_folder_emails', side_effect=fake_fetch):
                    result = web_outlook_app.fetch_account_emails(account, 'all', 0, 20)

        self.assertTrue(result['success'])
        self.assertEqual([email['folder'] for email in result['emails']], ['junkemail', 'inbox'])
        self.assertEqual(result['method'], 'method-inbox / method-junkemail')


class DatabaseConcurrencyTests(unittest.TestCase):
    def test_autocommit_connection_releases_write_lock_immediately(self):
        with tempfile.TemporaryDirectory(prefix='outlookEmail-lock-test-') as temp_dir:
            database_path = os.path.join(temp_dir, 'lock-test.db')
            with patch.object(web_outlook_app, 'DATABASE', database_path):
                setup = web_outlook_app.connect_database()
                setup.execute('CREATE TABLE events (value TEXT NOT NULL)')
                setup.commit()
                setup.close()

                background = web_outlook_app.connect_database(autocommit=True)
                foreground = web_outlook_app.connect_database()
                try:
                    background.execute("INSERT INTO events (value) VALUES ('background')")
                    foreground.execute("INSERT INTO events (value) VALUES ('foreground')")
                    foreground.commit()

                    self.assertEqual(
                        foreground.execute('SELECT COUNT(*) FROM events').fetchone()[0],
                        2,
                    )
                    self.assertEqual(
                        foreground.execute('PRAGMA busy_timeout').fetchone()[0],
                        web_outlook_app.SQLITE_BUSY_TIMEOUT_MS,
                    )
                finally:
                    foreground.close()
                    background.close()

    def test_forwarding_job_uses_autocommit_connection(self):
        connection = MagicMock(spec=sqlite3.Connection)
        with patch.object(web_outlook_app, 'connect_database', return_value=connection) as connect:
            with patch.object(web_outlook_app, 'get_forward_channels', return_value=[]):
                web_outlook_app.process_forwarding_job()

        self.assertIn(call(autocommit=True), connect.call_args_list)
        connection.close.assert_called()


class ForwardingTimeSafetyTests(unittest.TestCase):
    def test_parse_rfc_datetime_starting_with_t(self):
        parsed = web_outlook_app.parse_email_datetime('Thu, 15 Oct 2026 21:27:05 +0000 (UTC)')

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.isoformat(), '2026-10-15T21:27:05')

    def test_old_and_unparseable_email_dates_are_not_forwarded(self):
        connection = MagicMock(spec=sqlite3.Connection)
        connection.execute.return_value.fetchall.return_value = [{
            'id': 7,
            'email': 'user@outlook.com',
            'forward_last_checked_at': '2026-10-25T12:48:01Z',
            'password': '',
            'refresh_token': '',
            'imap_password': '',
        }]
        candidates = [
            {
                'id': 'old-rfc-message',
                'date': 'Thu, 15 Oct 2026 21:27:05 +0000 (UTC)',
                'folder': 'inbox',
                'mail_method': 'imap_new',
            },
            {
                'id': 'invalid-date-message',
                'date': '未知时间',
                'folder': 'inbox',
                'mail_method': 'imap_new',
            },
        ]

        with patch.object(web_outlook_app, 'connect_database', return_value=connection), \
                patch.object(web_outlook_app, 'get_forward_channels', return_value=['telegram']), \
                patch.object(web_outlook_app, 'get_setting_decrypted', return_value='bot-token'), \
                patch.object(web_outlook_app, 'get_setting', side_effect=lambda key, default='': {
                    'telegram_chat_id': 'chat-id',
                    'forward_email_window_minutes': '0',
                }.get(key, default)), \
                patch.object(web_outlook_app, 'get_bool_setting', return_value=False), \
                patch.object(web_outlook_app, 'fetch_forward_candidates', return_value={
                    'success': True,
                    'emails': candidates,
                }), \
                patch.object(web_outlook_app, 'fetch_forward_detail') as fetch_detail, \
                patch.object(web_outlook_app, 'send_forward_telegram') as send_telegram:
            web_outlook_app.process_forwarding_job()

        fetch_detail.assert_not_called()
        send_telegram.assert_not_called()

    def test_missing_cursor_is_initialized_without_scanning_existing_mail(self):
        connection = MagicMock(spec=sqlite3.Connection)
        connection.execute.return_value.fetchall.return_value = [{
            'id': 7,
            'email': 'user@outlook.com',
            'forward_last_checked_at': None,
            'password': '',
            'refresh_token': '',
            'imap_password': '',
        }]

        with patch.object(web_outlook_app, 'connect_database', return_value=connection), \
                patch.object(web_outlook_app, 'get_forward_channels', return_value=['telegram']), \
                patch.object(web_outlook_app, 'get_setting_decrypted', return_value='bot-token'), \
                patch.object(web_outlook_app, 'get_setting', side_effect=lambda key, default='': {
                    'telegram_chat_id': 'chat-id',
                    'forward_email_window_minutes': '0',
                }.get(key, default)), \
                patch.object(web_outlook_app, 'get_bool_setting', return_value=False), \
                patch.object(web_outlook_app, 'fetch_forward_candidates') as fetch_candidates:
            web_outlook_app.process_forwarding_job()

        fetch_candidates.assert_not_called()
        cursor_updates = [
            call_args for call_args in connection.execute.call_args_list
            if 'SET forward_last_checked_at = ?' in call_args.args[0]
        ]
        self.assertEqual(len(cursor_updates), 1)
        self.assertEqual(cursor_updates[0].args[1][1], 7)

    def test_email_after_cursor_is_forwarded(self):
        connection = MagicMock(spec=sqlite3.Connection)
        connection.execute.return_value.fetchall.return_value = [{
            'id': 7,
            'email': 'user@outlook.com',
            'forward_last_checked_at': '2026-10-25T12:48:01Z',
            'password': '',
            'refresh_token': '',
            'imap_password': '',
        }]
        candidate = {
            'id': 'new-message',
            'date': '2026-10-25T12:49:01Z',
            'folder': 'inbox',
            'mail_method': 'graph',
        }
        detail = {
            'id': 'new-message',
            'subject': 'New mail',
            'from': 'sender@example.com',
            'date': '2026-10-25T12:49:01Z',
            'body': 'Body',
            'body_type': 'text',
        }

        with patch.object(web_outlook_app, 'connect_database', return_value=connection), \
                patch.object(web_outlook_app, 'get_forward_channels', return_value=['telegram']), \
                patch.object(web_outlook_app, 'get_setting_decrypted', return_value='bot-token'), \
                patch.object(web_outlook_app, 'get_setting', side_effect=lambda key, default='': {
                    'telegram_chat_id': 'chat-id',
                    'forward_email_window_minutes': '0',
                }.get(key, default)), \
                patch.object(web_outlook_app, 'get_bool_setting', return_value=False), \
                patch.object(web_outlook_app, 'fetch_forward_candidates', return_value={
                    'success': True,
                    'emails': [candidate],
                }), \
                patch.object(web_outlook_app, 'fetch_forward_detail', return_value=detail), \
                patch.object(web_outlook_app, 'has_forward_log', return_value=False), \
                patch.object(web_outlook_app, 'send_forward_telegram', return_value=True) as send_telegram:
            web_outlook_app.process_forwarding_job()

        send_telegram.assert_called_once()


class ForwardingProtocolTests(unittest.TestCase):
    def setUp(self):
        self.account = {
            'id': 7,
            'email': 'user@outlook.com',
            'client_id': 'client-id',
            'refresh_token': 'refresh-token',
            'account_type': 'outlook',
            'mail_protocol': 'auto',
            'oauth_authority': 'common',
        }

    def test_imap_list_result_uses_same_imap_server_for_detail(self):
        with patch.object(web_outlook_app, 'get_account_proxy_url', return_value=''), \
                patch.object(web_outlook_app, 'get_account_proxy_failover_urls', return_value=[]), \
                patch.object(web_outlook_app, 'get_email_detail_graph') as graph, \
                patch.object(web_outlook_app, 'get_email_detail_imap', return_value={'id': '42'}) as imap:
            detail = web_outlook_app.fetch_forward_detail(
                self.account, '42', 'inbox', 'imap_old'
            )

        self.assertEqual(detail, {'id': '42'})
        graph.assert_not_called()
        self.assertEqual(imap.call_args.args[-1], web_outlook_app.IMAP_SERVER_OLD)

    def test_graph_list_result_uses_graph_for_detail(self):
        graph_detail = {
            'id': 'graph-id',
            'subject': 'Subject',
            'from': {'emailAddress': {'address': 'sender@example.com'}},
            'toRecipients': [],
            'ccRecipients': [],
            'body': {'content': 'Body', 'contentType': 'text'},
        }
        with patch.object(web_outlook_app, 'get_account_proxy_url', return_value=''), \
                patch.object(web_outlook_app, 'get_account_proxy_failover_urls', return_value=[]), \
                patch.object(web_outlook_app, 'get_email_detail_graph', return_value=graph_detail) as graph, \
                patch.object(web_outlook_app, 'get_email_detail_imap') as imap:
            detail = web_outlook_app.fetch_forward_detail(
                self.account, 'graph-id', 'inbox', 'graph'
            )

        self.assertEqual(detail['id'], 'graph-id')
        graph.assert_called_once()
        imap.assert_not_called()

    def test_only_explicit_expired_grant_is_classified_invalid(self):
        expired = {
            'code': 'GRAPH_TOKEN_FAILED',
            'details': '{"error":"invalid_grant","error_description":"refresh token has expired"}',
        }
        network = {
            'code': 'GRAPH_TOKEN_EXCEPTION',
            'type': 'ReadTimeout',
            'details': 'HTTPSConnectionPool timed out',
        }
        unauthorized_scope = {
            'code': 'IMAP_TOKEN_FAILED',
            'details': (
                '{"error":"invalid_grant","error_description":'
                '"one or more scopes requested are unauthorized or expired"}'
            ),
        }

        self.assertTrue(web_outlook_app.is_expired_token_error(expired))
        self.assertFalse(web_outlook_app.is_expired_token_error(network))
        self.assertFalse(web_outlook_app.is_expired_token_error(unauthorized_scope))

    def test_expired_grant_marks_account_and_network_error_does_not(self):
        expired = {
            'code': 'IMAP_TOKEN_FAILED',
            'details': '{"error":"invalid_grant","error_description":"grant was revoked"}',
        }
        with patch.object(web_outlook_app, 'fetch_account_folder_emails', return_value={
            'success': False,
            'error': '无法获取邮件，所有方式均失败',
            'details': {
                'graph': {
                    'code': 'GRAPH_TOKEN_FAILED',
                    'details': '{"error":"invalid_grant","error_description":"grant is expired"}',
                },
                'imap_new': expired,
                'imap_old': expired,
            },
        }):
            result = web_outlook_app.fetch_forward_candidates(self.account)

        self.assertTrue(result['token_invalid'])
        connection = MagicMock(spec=sqlite3.Connection)
        web_outlook_app.mark_account_token_invalid(connection, self.account['id'], result['error'])
        update_call = connection.execute.call_args
        self.assertIn("health_status = 'token_invalid'", update_call.args[0])
        self.assertEqual(update_call.args[1][1], self.account['id'])

        with patch.object(web_outlook_app, 'fetch_account_folder_emails', return_value={
            'success': False,
            'error': {'code': 'GRAPH_TOKEN_EXCEPTION', 'type': 'ConnectionResetError'},
            'details': {
                'graph': {'code': 'EMAIL_FETCH_FAILED', 'type': 'ReadTimeout'},
                'imap_new': {
                    'code': 'IMAP_TOKEN_FAILED',
                    'details': 'one or more scopes requested are unauthorized or expired',
                },
                'imap_old': {
                    'code': 'IMAP_TOKEN_FAILED',
                    'details': 'one or more scopes requested are unauthorized or expired',
                },
            },
        }):
            network_result = web_outlook_app.fetch_forward_candidates(self.account)
        self.assertFalse(network_result['token_invalid'])

    def test_auto_protocol_requires_all_token_paths_to_be_expired(self):
        expired = {
            'code': 'GRAPH_TOKEN_FAILED',
            'details': 'grant is expired; user must sign in again',
        }
        imap_scope_error = {
            'code': 'IMAP_TOKEN_FAILED',
            'details': 'one or more scopes requested are unauthorized or expired',
        }
        result = {
            'details': {
                'graph': {'code': 'EMAIL_FETCH_FAILED', 'type': 'ReadTimeout'},
                'imap_new': imap_scope_error,
                'imap_old': imap_scope_error,
            }
        }

        self.assertFalse(web_outlook_app.is_account_token_invalid_result(self.account, result))
        result['details'] = {
            'graph': expired,
            'imap_new': {'code': 'IMAP_TOKEN_FAILED', 'details': 'grant is expired'},
            'imap_old': {'code': 'IMAP_TOKEN_FAILED', 'details': 'grant is expired'},
        }
        self.assertTrue(web_outlook_app.is_account_token_invalid_result(self.account, result))

    def test_list_item_records_imap_old_as_detail_source(self):
        with patch.object(web_outlook_app, 'get_emails_graph', return_value={
            'success': False, 'error': 'graph failed'
        }), patch.object(web_outlook_app, 'get_emails_imap_with_server', side_effect=[
            {'success': False, 'error': 'new failed'},
            {'success': True, 'emails': [{'id': '42'}], 'has_more': False},
        ]):
            result = web_outlook_app.fetch_account_folder_emails(
                self.account, 'inbox', 0, 20
            )

        self.assertTrue(result['success'])
        self.assertEqual(result['method'], 'IMAP (Old)')
        self.assertEqual(result['emails'][0]['mail_method'], 'imap_old')

    def test_forwarding_query_skips_accounts_marked_token_invalid(self):
        connection = MagicMock(spec=sqlite3.Connection)
        connection.execute.return_value.fetchall.return_value = []
        with patch.object(web_outlook_app, 'connect_database', return_value=connection), \
                patch.object(web_outlook_app, 'get_forward_channels', return_value=['telegram']), \
                patch.object(web_outlook_app, 'get_setting_decrypted', return_value='bot-token'), \
                patch.object(web_outlook_app, 'get_setting', side_effect=lambda key, default='': {
                    'telegram_chat_id': 'chat-id',
                    'forward_include_junkemail': 'false',
                    'forward_email_window_minutes': '0',
                }.get(key, default)):
            web_outlook_app.process_forwarding_job()

        account_query = next(
            call.args[0] for call in connection.execute.call_args_list
            if 'SELECT * FROM accounts' in call.args[0]
        )
        self.assertIn("COALESCE(health_status, '') <> 'token_invalid'", account_query)

    def test_main_account_list_marks_expired_tokens(self):
        source_path = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), 'static', 'js', 'index', '02-groups.js'
        )
        with open(source_path, encoding='utf-8') as source_file:
            source = source_file.read()

        self.assertIn("acc.health_status === 'token_invalid'", source)
        self.assertIn('Token 过期', source)


class PublicApiRemovalTests(unittest.TestCase):
    def setUp(self):
        web_outlook_app.app.config['TESTING'] = True
        self.client = web_outlook_app.app.test_client()

    def test_external_account_and_mail_routes_are_not_registered(self):
        registered_paths = {rule.rule for rule in web_outlook_app.app.url_map.iter_rules()}
        for path in ('/api/external/accounts', '/api/external/emails'):
            with self.subTest(path=path):
                self.assertNotIn(path, registered_paths)


class BatchForwardingApiTests(unittest.TestCase):
    def setUp(self):
        self.app = web_outlook_app.app
        self.app.config['TESTING'] = True
        self.app.config['WTF_CSRF_ENABLED'] = False
        self.client = self.app.test_client()

        with self.client.session_transaction() as sess:
            sess['logged_in'] = True

        with self.app.app_context():
            db = web_outlook_app.get_db()
            db.execute('DELETE FROM account_tags')
            db.execute('DELETE FROM account_aliases')
            db.execute('DELETE FROM account_refresh_logs')
            db.execute('DELETE FROM accounts')
            db.execute('DELETE FROM tags')
            db.execute("DELETE FROM groups WHERE name NOT IN ('默认分组', '临时邮箱')")
            db.commit()

            self.assertTrue(web_outlook_app.add_account(
                'disabled-forward@outlook.com',
                'password123',
                'client-id-disabled',
                'refresh-token-disabled',
                group_id=1,
                forward_enabled=False
            ))
            self.assertTrue(web_outlook_app.add_account(
                'enabled-forward@outlook.com',
                'password123',
                'client-id-enabled',
                'refresh-token-enabled',
                group_id=1,
                forward_enabled=True
            ))

            disabled_account = web_outlook_app.get_account_by_email('disabled-forward@outlook.com')
            enabled_account = web_outlook_app.get_account_by_email('enabled-forward@outlook.com')
            self.assertIsNotNone(disabled_account)
            self.assertIsNotNone(enabled_account)

            self.disabled_account_id = disabled_account['id']
            self.enabled_account_id = enabled_account['id']
            self.disabled_cursor_before = disabled_account.get('forward_last_checked_at')
            self.enabled_cursor_before = enabled_account.get('forward_last_checked_at')

    def test_batch_enable_forwarding_only_updates_disabled_accounts(self):
        response = self.client.post(
            '/api/accounts/batch-update-forwarding',
            json={
                'account_ids': [self.disabled_account_id, self.enabled_account_id],
                'forward_enabled': True,
            }
        )

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])
        self.assertEqual(payload['updated_count'], 1)
        self.assertEqual(payload['unchanged_count'], 1)

        with self.app.app_context():
            disabled_account = web_outlook_app.get_account_by_id(self.disabled_account_id)
            enabled_account = web_outlook_app.get_account_by_id(self.enabled_account_id)

        self.assertTrue(disabled_account['forward_enabled'])
        self.assertTrue(disabled_account['forward_last_checked_at'])
        self.assertEqual(enabled_account['forward_last_checked_at'], self.enabled_cursor_before)

    def test_batch_disable_forwarding_only_updates_enabled_accounts(self):
        response = self.client.post(
            '/api/accounts/batch-update-forwarding',
            json={
                'account_ids': [self.disabled_account_id, self.enabled_account_id],
                'forward_enabled': False,
            }
        )

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])
        self.assertEqual(payload['updated_count'], 1)
        self.assertEqual(payload['unchanged_count'], 1)

        with self.app.app_context():
            disabled_account = web_outlook_app.get_account_by_id(self.disabled_account_id)
            enabled_account = web_outlook_app.get_account_by_id(self.enabled_account_id)

        self.assertFalse(disabled_account['forward_enabled'])
        self.assertEqual(disabled_account['forward_last_checked_at'], self.disabled_cursor_before)
        self.assertFalse(enabled_account['forward_enabled'])
        self.assertEqual(enabled_account['forward_last_checked_at'], self.enabled_cursor_before)


class AssetRenderingTests(unittest.TestCase):
    def setUp(self):
        self.app = web_outlook_app.app
        self.app.config['TESTING'] = True
        self.client = self.app.test_client()

    def test_index_uses_bundled_stylesheet_route(self):
        with self.client.session_transaction() as session:
            session['logged_in'] = True

        response = self.client.get('/')
        html = response.get_data(as_text=True)

        self.assertEqual(response.status_code, 200)
        self.assertIn('href="/assets/index.css"', html)
        self.assertNotIn('href="/static/index.css"', html)

    def test_index_core_keeps_shared_helpers_after_refresh_ui_removal(self):
        response = self.client.get('/static/js/index/01-core.js')
        script = response.get_data(as_text=True)
        response.close()

        self.assertIn('function closeAllModals()', script)
        self.assertIn('function formatDateTime(dateStr)', script)
        self.assertNotIn("getElementById('refreshProgress')", script)

    def test_bundled_stylesheet_contains_combined_css_without_imports(self):
        response = self.client.get('/assets/index.css')
        css = response.get_data(as_text=True)

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, 'text/css')
        self.assertNotIn('@import', css)
        self.assertIn('.toast', css)
        self.assertIn('.group-panel', css)
        self.assertIn('.account-panel', css)

    def test_account_pool_ui_and_actions_are_present(self):
        with self.client.session_transaction() as session:
            session['logged_in'] = True

        index_html = self.client.get('/').get_data(as_text=True)
        response = self.client.get('/account-pool')
        html = response.get_data(as_text=True)
        script_response = self.client.get('/static/js/account-pool.js')
        script = script_response.get_data(as_text=True)
        script_response.close()

        self.assertEqual(response.status_code, 200)
        self.assertNotIn('id="accountPoolPage"', index_html)
        self.assertNotIn('/static/js/account-pool.js', index_html)
        self.assertIn('href="/account-pool"', index_html)
        self.assertIn('id="pageSize"', html)
        self.assertIn('id="domainFilter"', html)
        self.assertIn('id="testAllBtn"', html)
        self.assertIn('id="retryFailedBtn"', html)
        self.assertNotIn('id="refreshModal"', html)
        self.assertNotIn('Token刷新管理', html)
        self.assertIn('/static/js/account-pool.js', html)
        self.assertIn('/api/accounts/${id}/refresh', script)
        self.assertIn('/api/accounts/refresh-selected', script)
        self.assertIn("fetch('/api/csrf-token')", script)
        self.assertIn("'X-CSRFToken': csrfToken", script)
        self.assertIn("${a.email || ''}----${a.password || ''}", script)
        self.assertIn("----${a.client_id || ''}----${a.refresh_token || ''}", script)

    def test_removed_api_and_temp_mail_settings_are_not_exposed(self):
        with self.client.session_transaction() as session:
            session['logged_in'] = True

        html = self.client.get('/').get_data(as_text=True)
        settings = self.client.get('/api/settings').get_json()['settings']

        self.assertNotIn('settingsApiKey', html)
        self.assertNotIn('settingsDuckmailBaseUrl', html)
        self.assertNotIn('settingsCloudflareWorkerDomain', html)
        for key in ('external_api_key', 'gptmail_api_key', 'duckmail_base_url', 'duckmail_api_key',
                    'cloudflare_worker_domain', 'cloudflare_email_domains', 'cloudflare_admin_password'):
            self.assertNotIn(key, settings)

    def test_account_pool_csrf_header_allows_post_request(self):
        with self.client.session_transaction() as session:
            session['logged_in'] = True

        token = self.client.get('/api/csrf-token').get_json()['csrf_token']
        response = self.client.post('/api/accounts/999999/refresh', headers={'X-CSRFToken': token})

        self.assertEqual(response.status_code, 200)
        self.assertFalse(response.get_json()['success'])


class RefreshTokenProxyFallbackTests(unittest.TestCase):
    def setUp(self):
        self.app = web_outlook_app.app
        self.app.config['TESTING'] = True
        self.client = self.app.test_client()

        with self.client.session_transaction() as sess:
            sess['logged_in'] = True

        with self.app.app_context():
            db = web_outlook_app.get_db()
            db.execute('DELETE FROM account_refresh_logs')
            db.execute('DELETE FROM accounts')
            db.execute("DELETE FROM groups WHERE name NOT IN ('默认分组', '临时邮箱')")
            db.commit()

            group_id = web_outlook_app.add_group(
                '代理刷新组',
                '测试代理失败后直连重试',
                '#225588',
                'socks5://127.0.0.1:1080',
                'http://127.0.0.1:7891',
                'direct',
            )
            self.assertIsNotNone(group_id)
            self.group_id = group_id

            added = web_outlook_app.add_account(
                'proxy-refresh@outlook.com',
                'password123',
                '24d9a0ed-8787-4584-883c-2fd79308940a',
                '0.AXEA_refresh',
                group_id=group_id,
                remark='代理刷新测试账号',
                forward_enabled=False,
            )
            self.assertTrue(added)

            account = web_outlook_app.get_account_by_email('proxy-refresh@outlook.com')
            self.assertIsNotNone(account)
            self.account_id = account['id']

    def test_refresh_account_retries_with_fallback_proxies_in_order(self):
        class FakeResponse:
            status_code = 200

            @staticmethod
            def json():
                return {'access_token': 'access-token'}

        proxy_error = web_outlook_app.requests.exceptions.ConnectionError(
            'SOCKSHTTPSConnectionPool(host=\'login.microsoftonline.com\', port=443): '
            'Max retries exceeded with url: /common/oauth2/v2.0/token '
            '(Caused by NewConnectionError("SOCKSHTTPSConnection(host=\'login.microsoftonline.com\', '
            'port=443): Failed to establish a new connection: 0x04: Host unreachable"))'
        )

        http_proxy_error = web_outlook_app.requests.exceptions.ProxyError(
            'HTTPSConnectionPool(host=\'login.microsoftonline.com\', port=443): '
            'Max retries exceeded with url: /common/oauth2/v2.0/token '
            '(Caused by ProxyError(\'Unable to connect to proxy\', OSError(\'proxy connect failed\')))'
        )

        with patch.object(
            web_outlook_app.requests,
            'request',
            side_effect=[proxy_error, http_proxy_error, FakeResponse(), FakeResponse()],
        ) as mocked_request, patch.object(
            web_outlook_app,
            'get_access_token_imap_result',
            return_value={'success': False, 'error': 'IMAP scope unavailable'},
        ):
            response = self.client.post(f'/api/accounts/{self.account_id}/refresh')

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])
        self.assertEqual(payload['message'], '账号权限测试通过')
        self.assertEqual(payload['permissions']['health_status'], 'graph_only')
        self.assertEqual(mocked_request.call_count, 4)
        self.assertEqual(
            mocked_request.call_args_list[0].kwargs['proxies'],
            {'http': 'socks5://127.0.0.1:1080', 'https': 'socks5://127.0.0.1:1080'},
        )
        self.assertEqual(
            mocked_request.call_args_list[1].kwargs['proxies'],
            {'http': 'http://127.0.0.1:7891', 'https': 'http://127.0.0.1:7891'},
        )
        self.assertEqual(
            mocked_request.call_args_list[2].kwargs['proxies'],
            {'http': None, 'https': None, 'all': None},
        )

        with self.app.app_context():
            db = web_outlook_app.get_db()
            log_row = db.execute(
                '''
                SELECT status, error_message
                FROM account_refresh_logs
                WHERE account_id = ?
                ORDER BY id DESC
                LIMIT 1
                ''',
                (self.account_id,),
            ).fetchone()

        self.assertIsNotNone(log_row)
        self.assertEqual(log_row['status'], 'success')
        self.assertIsNone(log_row['error_message'])

    def test_group_api_persists_proxy_failover_fields(self):
        response = self.client.put(
            f'/api/groups/{self.group_id}',
            json={
                'name': '代理刷新组',
                'description': '测试代理失败后直连重试',
                'color': '#225588',
                'proxy_url': 'socks5://127.0.0.1:1080',
                'fallback_proxy_url_1': 'socks5://127.0.0.1:2080',
                'fallback_proxy_url_2': '直连',
                'sort_position': 1,
            }
        )

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])

        response = self.client.get(f'/api/groups/{self.group_id}')
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])
        self.assertEqual(payload['group']['fallback_proxy_url_1'], 'socks5://127.0.0.1:2080')
        self.assertEqual(payload['group']['fallback_proxy_url_2'], '直连')


if __name__ == '__main__':
    unittest.main()
