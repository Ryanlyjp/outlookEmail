from __future__ import annotations

from typing import TYPE_CHECKING, Any, Dict, Optional

if TYPE_CHECKING:
    # These segmented files are executed into the shared `web_outlook_app`
    # globals at runtime. Importing from the assembled module keeps IDE
    # inspections from flagging the shared names as unresolved.
    from web_outlook_app import *  # noqa: F403


# ==================== 定时任务调度器 ====================

def get_bool_setting(key: str, default: bool = False) -> bool:
    value = str(get_setting(key, 'true' if default else 'false')).strip().lower()
    return value in ('1', 'true', 'yes', 'on')


def normalize_forward_channel_settings(raw_channels: Any) -> list[str]:
    if isinstance(raw_channels, str):
        values = raw_channels.split(',')
    elif isinstance(raw_channels, (list, tuple, set)):
        values = list(raw_channels)
    else:
        values = []

    normalized = []
    channel_aliases = {
        'email': FORWARD_CHANNEL_SMTP_SETTING,
        'smtp': FORWARD_CHANNEL_SMTP_SETTING,
        'tg': FORWARD_CHANNEL_TG_SETTING,
        'telegram': FORWARD_CHANNEL_TG_SETTING,
    }

    for item in values:
        channel = channel_aliases.get(str(item).strip().lower())
        if channel and channel not in normalized:
            normalized.append(channel)
    return normalized


def normalize_smtp_forward_provider(value: str) -> str:
    provider = str(value or '').strip().lower()
    if provider not in SMTP_FORWARD_PROVIDERS:
        return 'custom'
    return provider


def get_configured_forward_channels() -> list[str]:
    channels = []
    if get_setting('email_forward_recipient', '').strip() and get_setting('smtp_host', '').strip():
        channels.append(FORWARD_CHANNEL_SMTP_SETTING)
    if get_setting_decrypted('telegram_bot_token', '').strip() and get_setting('telegram_chat_id', '').strip():
        channels.append(FORWARD_CHANNEL_TG_SETTING)
    return channels


def get_forward_channels() -> list[str]:
    raw_channels = get_setting('forward_channels', 'auto').strip().lower()
    if raw_channels in ('', 'auto'):
        return get_configured_forward_channels()
    if raw_channels == 'none':
        return []
    return normalize_forward_channel_settings(raw_channels)


def has_forward_log(conn, account_id: int, message_id: str, channel: str) -> bool:
    row = conn.execute(
        'SELECT 1 FROM forward_logs WHERE account_id = ? AND message_id = ? AND channel = ? LIMIT 1',
        (account_id, str(message_id), channel)
    ).fetchone()
    return row is not None


def record_forward_log(conn, account_id: int, message_id: str, channel: str):
    conn.execute(
        'INSERT OR IGNORE INTO forward_logs (account_id, message_id, channel) VALUES (?, ?, ?)',
        (account_id, str(message_id), channel)
    )


def stringify_forward_error(error: Any) -> str:
    if error is None:
        return ''
    if isinstance(error, str):
        return sanitize_error_details(error)[:500]
    if isinstance(error, dict):
        parts = []
        message = str(error.get('message') or error.get('error') or '').strip()
        code = str(error.get('code') or '').strip()
        err_type = str(error.get('type') or '').strip()
        details = error.get('details')
        if message:
            parts.append(message)
        if code:
            parts.append(f'code={code}')
        if err_type:
            parts.append(f'type={err_type}')
        if details:
            details_text = details if isinstance(details, str) else json.dumps(details, ensure_ascii=True)
            if details_text and details_text not in parts:
                parts.append(str(details_text))
        return sanitize_error_details(' | '.join(parts) or str(error))[:500]
    return sanitize_error_details(str(error))[:500]


def is_expired_token_error(error: Any) -> bool:
    """Match only explicit OAuth grant expiry/revocation responses."""
    if not error:
        return False
    if isinstance(error, dict):
        error_code = str(error.get('code') or '')
        if error_code not in ('GRAPH_TOKEN_FAILED', 'IMAP_TOKEN_FAILED'):
            return any(is_expired_token_error(value) for value in error.values())
        text = ' '.join(str(error.get(key) or '') for key in ('message', 'details')).lower()
    elif isinstance(error, (list, tuple)):
        return any(is_expired_token_error(value) for value in error)
    else:
        text = str(error).lower()
        if 'graph_token_failed' not in text and 'imap_token_failed' not in text:
            return False
    return any(marker in text for marker in (
        'refresh token has expired',
        'refresh token is expired',
        'refresh token was revoked',
        'refresh token has been revoked',
        'grant is expired',
        'grant has expired',
        'grant was revoked',
    ))


def is_account_token_invalid_result(account: Dict[str, Any], result: Dict[str, Any]) -> bool:
    if account.get('account_type') == 'imap':
        return False

    details = result.get('details')
    if not isinstance(details, dict):
        return False

    protocol = normalize_mail_protocol(account.get('mail_protocol'))
    if protocol == 'graph':
        return is_expired_token_error(details.get('graph'))
    if protocol == 'imap_oauth':
        return all(is_expired_token_error(details.get(key)) for key in ('imap_new', 'imap_old'))
    return all(is_expired_token_error(details.get(key)) for key in ('graph', 'imap_new', 'imap_old'))


def mark_account_token_invalid(conn, account_id: int, error: Any) -> None:
    conn.execute(
        '''
        UPDATE accounts
        SET health_status = 'token_invalid', health_detail = ?,
            graph_status = 'fail', imap_status = 'fail',
            last_protocol_test_at = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        ''',
        (stringify_forward_error(error), account_id),
    )


def send_forward_email(subject: str, body_text: str, body_html: str = '') -> bool:
    recipient = get_setting('email_forward_recipient', '').strip()
    host = get_setting('smtp_host', '').strip()
    if not recipient or not host:
        return False

    port = int(get_setting('smtp_port', '465') or 465)
    username = get_setting('smtp_username', '').strip()
    password = get_setting_decrypted('smtp_password', '').strip()
    smtp_from_email = get_setting('smtp_from_email', '').strip()
    from_email = smtp_from_email or username
    if not from_email:
        return False
    use_tls = get_bool_setting('smtp_use_tls', False)
    use_ssl = get_bool_setting('smtp_use_ssl', True)

    message = EmailMessage()
    message['From'] = from_email
    message['To'] = recipient
    message['Subject'] = subject
    message.set_content(body_text)
    if body_html:
        message.add_alternative(body_html, subtype='html')

    smtp_cls = smtplib.SMTP_SSL if use_ssl else smtplib.SMTP
    with smtp_cls(host, port, timeout=20) as client:
        if not use_ssl:
            client.ehlo()
            if use_tls:
                client.starttls()
                client.ehlo()
        if username:
            client.login(username, password)
        client.send_message(message)
    return True


def send_forward_email_with_config(config: Dict[str, Any], subject: str, body_text: str, body_html: str = '') -> bool:
    recipient = str(config.get('recipient', '') or '').strip()
    host = str(config.get('host', '') or '').strip()
    if not recipient or not host:
        return False

    port = int(config.get('port') or 465)
    username = str(config.get('username', '') or '').strip()
    password = str(config.get('password', '') or '')
    from_email = str(config.get('from_email', '') or '').strip() or username
    if not from_email:
        return False
    use_tls = str(config.get('use_tls', '')).lower() in ('1', 'true', 'yes', 'on')
    use_ssl = str(config.get('use_ssl', '')).lower() in ('1', 'true', 'yes', 'on')

    message = EmailMessage()
    message['From'] = from_email
    message['To'] = recipient
    message['Subject'] = subject
    message.set_content(body_text)
    if body_html:
        message.add_alternative(body_html, subtype='html')

    smtp_cls = smtplib.SMTP_SSL if use_ssl else smtplib.SMTP
    with smtp_cls(host, port, timeout=20) as client:
        if not use_ssl:
            client.ehlo()
            if use_tls:
                client.starttls()
                client.ehlo()
        if username:
            client.login(username, password)
        client.send_message(message)
    return True


def send_forward_telegram(text: str) -> bool:
    bot_token = get_setting_decrypted('telegram_bot_token', '').strip()
    chat_id = get_setting('telegram_chat_id', '').strip()
    if not bot_token or not chat_id:
        return False
    response = requests.post(
        f'https://api.telegram.org/bot{bot_token}/sendMessage',
        json={
            'chat_id': chat_id,
            'text': text[:4000],
            'disable_web_page_preview': True,
        },
        timeout=15
    )
    return response.ok


def send_forward_telegram_with_config(config: Dict[str, Any], text: str) -> bool:
    bot_token = str(config.get('bot_token', '') or '').strip()
    chat_id = str(config.get('chat_id', '') or '').strip()
    if not bot_token or not chat_id:
        return False
    response = requests.post(
        f'https://api.telegram.org/bot{bot_token}/sendMessage',
        json={
            'chat_id': chat_id,
            'text': text[:4000],
            'disable_web_page_preview': True,
        },
        timeout=15
    )
    return response.ok


def build_forward_payload(account: Dict[str, Any], email_detail: Dict[str, Any]) -> tuple[str, str, str, str]:
    subject = email_detail.get('subject') or '无主题'
    sender = email_detail.get('from') or '未知'
    received_at = email_detail.get('date') or ''
    body = email_detail.get('body') or ''
    body_text = strip_html_content(body) if email_detail.get('body_type') == 'html' else strip_html_content(body.replace('<br>', '\n'))
    body_text = body_text[:2000]

    title = f"[邮件转发] {subject}"
    plain = f"账号: {account.get('email','')}\n发件人: {sender}\n时间: {received_at}\n主题: {subject}\n\n{body_text}"
    html_body = (
        f"<p><strong>账号:</strong> {html.escape(account.get('email', ''))}</p>"
        f"<p><strong>发件人:</strong> {html.escape(sender)}</p>"
        f"<p><strong>时间:</strong> {html.escape(received_at)}</p>"
        f"<p><strong>主题:</strong> {html.escape(subject)}</p><hr>{body}"
    )
    telegram_text = f"新邮件转发\n账号: {account.get('email','')}\n发件人: {sender}\n主题: {subject}\n时间: {received_at}\n\n{body_text[:1200]}"
    return title, plain, html_body, telegram_text


def fetch_forward_candidates(account: Dict[str, Any], top: int = 20, folder: str = 'inbox') -> Dict[str, Any]:
    proxy_url = get_account_proxy_url(account)
    fallback_proxy_urls = get_account_proxy_failover_urls(account)
    result = fetch_account_folder_emails(account, folder, 0, top, proxy_url, fallback_proxy_urls)
    if not result.get('success'):
        token_invalid = is_account_token_invalid_result(account, result)
        error_source = result.get('details') if token_invalid else result.get('error') or result.get('details')
        return {
            'success': False,
            'emails': [],
            'error': stringify_forward_error(error_source or '获取邮件失败'),
            'token_invalid': token_invalid,
        }
    return {
        'success': True,
        'emails': [
            dict(item, mail_method=item.get('mail_method') or result.get('method', ''))
            for item in result.get('emails', [])
        ],
        'error': '',
    }


def build_forward_cursor_reset(account: Dict[str, Any], mode: str = 'window', lookback_minutes: Optional[int] = None) -> tuple[Optional[str], str, int]:
    normalized_mode = str(mode or 'window').strip().lower()

    if lookback_minutes is None:
        try:
            lookback_minutes = max(0, min(10080, int(get_setting('forward_email_window_minutes', '0') or '0')))
        except (TypeError, ValueError):
            lookback_minutes = 0
    else:
        lookback_minutes = max(0, min(10080, int(lookback_minutes or 0)))

    if normalized_mode == 'clear':
        return None, '已清空转发游标，下次会从当前可拉取到的最近邮件重新扫描', lookback_minutes

    if lookback_minutes > 0:
        cursor_value = (datetime.now(timezone.utc) - timedelta(minutes=lookback_minutes)).isoformat()
        return cursor_value, f'已回退转发游标，接下来会重扫最近 {lookback_minutes} 分钟内的邮件', lookback_minutes

    return None, '当前未限制转发时间范围，已清空转发游标并准备重扫最近邮件', lookback_minutes


def fetch_forward_detail(account: Dict[str, Any], message_id: str, folder: str = 'inbox',
                         mail_method: str = 'graph') -> Optional[Dict[str, Any]]:
    return fetch_account_email_detail(account, message_id, folder, mail_method)


def process_forwarding_job():
    with app.app_context():
        # Forwarding performs slow external requests between database writes.
        # Autocommit prevents those requests from extending a SQLite write lock.
        conn = connect_database(autocommit=True)
        try:
            forward_channels = set(get_forward_channels())
            email_enabled = FORWARD_CHANNEL_SMTP_SETTING in forward_channels and bool(
                get_setting('email_forward_recipient', '').strip() and get_setting('smtp_host', '').strip()
            )
            telegram_enabled = FORWARD_CHANNEL_TG_SETTING in forward_channels and bool(
                get_setting_decrypted('telegram_bot_token', '').strip() and get_setting('telegram_chat_id', '').strip()
            )
            include_junkemail = get_bool_setting('forward_include_junkemail', False)
            try:
                forward_window_minutes = max(0, min(10080, int(get_setting('forward_email_window_minutes', '0') or '0')))
            except (TypeError, ValueError):
                forward_window_minutes = 0
            forward_window_start = datetime.now() - timedelta(minutes=forward_window_minutes) if forward_window_minutes > 0 else None
            if not email_enabled and not telegram_enabled:
                print('[forward] skip job: no active channels configured')
                return

            accounts = conn.execute(
                """SELECT * FROM accounts
                   WHERE status = 'active' AND forward_enabled = 1
                     AND COALESCE(health_status, '') <> 'token_invalid'"""
            ).fetchall()
            print(
                f"[forward] start job: accounts={len(accounts)} email_enabled={email_enabled} telegram_enabled={telegram_enabled}"
            )
            for row in accounts:
                account = dict(row)
                if account.get('password'):
                    try:
                        account['password'] = decrypt_data(account['password'])
                    except Exception:
                        pass
                if account.get('refresh_token'):
                    try:
                        account['refresh_token'] = decrypt_data(account['refresh_token'])
                    except Exception:
                        pass
                if account.get('imap_password'):
                    try:
                        account['imap_password'] = decrypt_data(account['imap_password'])
                    except Exception:
                        pass

                cursor_time = parse_email_datetime(account.get('forward_last_checked_at', ''))
                folders_to_scan = ['inbox']
                if include_junkemail:
                    folders_to_scan.append('junkemail')
                emails = []
                had_processing_failure = False
                token_invalid = False
                for folder_name in folders_to_scan:
                    folder_result = fetch_forward_candidates(account, 20, folder_name)
                    if not folder_result.get('success'):
                        had_processing_failure = True
                        error_message = f'{folder_name} 候选邮件拉取失败: {folder_result.get("error") or "未知错误"}'
                        log_forwarding_result(
                            account['id'],
                            account.get('email', ''),
                            f'folder:{folder_name}',
                            'fetch_candidates',
                            'failed',
                            error_message,
                            db_conn=conn,
                        )
                        app.logger.warning(
                            '[forward] fetch candidates failed: account=%s folder=%s error=%s',
                            account.get('email', ''),
                            folder_name,
                            folder_result.get('error') or 'unknown',
                        )
                        if folder_result.get('token_invalid'):
                            mark_account_token_invalid(conn, account['id'], folder_result.get('error'))
                            token_invalid = True
                            app.logger.warning(
                                '[forward] stop retrying invalid token: account=%s',
                                account.get('email', ''),
                            )
                            break
                        continue
                    emails.extend(folder_result.get('emails', []))
                if token_invalid:
                    continue
                recent_emails = []
                skipped_before_cursor = 0
                email_success_count = 0
                telegram_success_count = 0
                latest_success_time = cursor_time
                for item in emails:
                    dt = parse_email_datetime(item.get('date', ''))
                    if forward_window_start and dt and dt < forward_window_start:
                        app.logger.info(
                            '[forward] skip email older than window: account=%s message_id=%s email_time=%s window_start=%s',
                            account.get('email', ''),
                            item.get('id', ''),
                            item.get('date', ''),
                            forward_window_start.isoformat(),
                        )
                        continue
                    if cursor_time and dt and dt <= cursor_time:
                        skipped_before_cursor += 1
                        app.logger.info(
                            '[forward] skip email before cursor: account=%s message_id=%s email_time=%s cursor=%s',
                            account.get('email', ''),
                            item.get('id', ''),
                            item.get('date', ''),
                            account.get('forward_last_checked_at', ''),
                        )
                        continue
                    recent_emails.append((dt, item))
                app.logger.info(
                    '[forward] account candidates: account=%s fetched=%s pending=%s skipped_before_cursor=%s cursor=%s',
                    account.get('email', ''),
                    len(emails),
                    len(recent_emails),
                    skipped_before_cursor,
                    account.get('forward_last_checked_at', ''),
                )

                recent_emails.sort(key=lambda pair: pair[0] or datetime.min)

                for item_time, item in recent_emails:
                    detail = fetch_forward_detail(
                        account,
                        item.get('id'),
                        item.get('folder', 'inbox'),
                        item.get('mail_method', 'graph'),
                    )
                    if not detail:
                        had_processing_failure = True
                        log_forwarding_result(
                            account['id'],
                            account.get('email', ''),
                            item.get('id', ''),
                            'detail',
                            'failed',
                            '获取邮件详情失败',
                            db_conn=conn,
                        )
                        app.logger.warning(
                            '[forward] skip email detail fetch failed: account=%s message_id=%s',
                            account.get('email', ''),
                            item.get('id', ''),
                        )
                        continue
                    title, plain, html_body, telegram_text = build_forward_payload(account, detail)
                    message_processed = False
                    message_failed = False
                    if email_enabled:
                        if has_forward_log(conn, account['id'], detail['id'], FORWARD_CHANNEL_EMAIL):
                            app.logger.info(
                                '[forward] skip already forwarded email: account=%s message_id=%s channel=%s',
                                account.get('email', ''),
                                detail.get('id', ''),
                                FORWARD_CHANNEL_EMAIL,
                            )
                            message_processed = True
                        else:
                            try:
                                if send_forward_email(title, plain, html_body):
                                    record_forward_log(conn, account['id'], detail['id'], FORWARD_CHANNEL_EMAIL)
                                    log_forwarding_result(
                                        account['id'],
                                        account.get('email', ''),
                                        detail.get('id', ''),
                                        FORWARD_CHANNEL_EMAIL,
                                        'success',
                                        db_conn=conn,
                                    )
                                    email_success_count += 1
                                    message_processed = True
                                else:
                                    message_failed = True
                                    log_forwarding_result(
                                        account['id'],
                                        account.get('email', ''),
                                        detail.get('id', ''),
                                        FORWARD_CHANNEL_EMAIL,
                                        'failed',
                                        'SMTP 转发返回失败',
                                        db_conn=conn,
                                    )
                                    app.logger.warning(
                                        '[forward] send email returned false: account=%s message_id=%s channel=%s',
                                        account.get('email', ''),
                                        detail.get('id', ''),
                                        FORWARD_CHANNEL_EMAIL,
                                    )
                            except Exception as exc:
                                message_failed = True
                                log_forwarding_result(
                                    account['id'],
                                    account.get('email', ''),
                                    detail.get('id', ''),
                                    FORWARD_CHANNEL_EMAIL,
                                    'failed',
                                    str(exc),
                                    db_conn=conn,
                                )
                                app.logger.warning(
                                    '[forward] send email failed: account=%s message_id=%s channel=%s error=%s',
                                    account.get('email', ''),
                                    detail.get('id', ''),
                                    FORWARD_CHANNEL_EMAIL,
                                    str(exc),
                                )
                    if telegram_enabled:
                        if has_forward_log(conn, account['id'], detail['id'], FORWARD_CHANNEL_TELEGRAM):
                            app.logger.info(
                                '[forward] skip already forwarded email: account=%s message_id=%s channel=%s',
                                account.get('email', ''),
                                detail.get('id', ''),
                                FORWARD_CHANNEL_TELEGRAM,
                            )
                            message_processed = True
                        else:
                            try:
                                if send_forward_telegram(telegram_text):
                                    record_forward_log(conn, account['id'], detail['id'], FORWARD_CHANNEL_TELEGRAM)
                                    log_forwarding_result(
                                        account['id'],
                                        account.get('email', ''),
                                        detail.get('id', ''),
                                        FORWARD_CHANNEL_TELEGRAM,
                                        'success',
                                        db_conn=conn,
                                    )
                                    telegram_success_count += 1
                                    message_processed = True
                                else:
                                    message_failed = True
                                    log_forwarding_result(
                                        account['id'],
                                        account.get('email', ''),
                                        detail.get('id', ''),
                                        FORWARD_CHANNEL_TELEGRAM,
                                        'failed',
                                        'Telegram 转发返回失败',
                                        db_conn=conn,
                                    )
                                    app.logger.warning(
                                        '[forward] send telegram returned false: account=%s message_id=%s channel=%s',
                                        account.get('email', ''),
                                        detail.get('id', ''),
                                        FORWARD_CHANNEL_TELEGRAM,
                                    )
                            except Exception as exc:
                                message_failed = True
                                log_forwarding_result(
                                    account['id'],
                                    account.get('email', ''),
                                    detail.get('id', ''),
                                    FORWARD_CHANNEL_TELEGRAM,
                                    'failed',
                                    str(exc),
                                    db_conn=conn,
                                )
                                app.logger.warning(
                                    '[forward] send telegram failed: account=%s message_id=%s channel=%s error=%s',
                                    account.get('email', ''),
                                    detail.get('id', ''),
                                    FORWARD_CHANNEL_TELEGRAM,
                                    str(exc),
                                )

                    if message_failed:
                        had_processing_failure = True
                        continue
                    if message_processed and item_time and (latest_success_time is None or item_time > latest_success_time):
                        latest_success_time = item_time

                cursor_value = account.get('forward_last_checked_at', '')
                cursor_updated = False
                if latest_success_time and (cursor_time is None or latest_success_time > cursor_time):
                    cursor_value = latest_success_time.isoformat()
                    conn.execute(
                        'UPDATE accounts SET forward_last_checked_at = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                        (cursor_value, account['id'])
                    )
                    cursor_updated = True
                else:
                    conn.execute(
                        'UPDATE accounts SET updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                        (account['id'],)
                    )
                conn.commit()
                print(
                    f"[forward] account done: account={account.get('email', '')} email_success={email_success_count} telegram_success={telegram_success_count} cursor_updated={cursor_updated} cursor={cursor_value} had_failure={had_processing_failure}"
                )
        except Exception as exc:
            print(f"[forward] job failed: {str(exc)}")
            raise
        finally:
            conn.close()


@app.route('/api/accounts/trigger-forwarding-check', methods=['POST'])
@login_required
def api_trigger_forwarding_check():
    """手动触发一次转发检查"""
    try:
        process_forwarding_job()
        return jsonify({'success': True, 'message': '已触发一次转发检查，请查看转发历史或容器日志'})
    except Exception as exc:
        return jsonify({'success': False, 'error': f'触发转发检查失败: {str(exc)}'})


@app.route('/api/accounts/<int:account_id>/forwarding/reset-cursor', methods=['POST'])
@login_required
def api_reset_account_forward_cursor(account_id):
    """回退或清空单个账号的转发游标，并可选触发一次重扫。"""
    account = get_account_by_id(account_id)
    if not account:
        return jsonify({'success': False, 'error': '账号不存在'}), 404

    data = request.json or {}
    mode = str(data.get('mode', 'window') or 'window')
    lookback_minutes = data.get('lookback_minutes')
    trigger_check = bool(data.get('trigger_check', True))

    try:
        cursor_value, reset_message, effective_lookback = build_forward_cursor_reset(account, mode, lookback_minutes)
    except (TypeError, ValueError):
        return jsonify({'success': False, 'error': '回退时间参数无效'}), 400

    if not set_account_forward_cursor(account_id, cursor_value):
        return jsonify({'success': False, 'error': '重置转发游标失败'}), 500

    triggered = False
    if trigger_check:
        process_forwarding_job()
        triggered = True

    action_message = reset_message
    if triggered:
        action_message += '，并已立即触发一次转发检查'

    return jsonify({
        'success': True,
        'message': action_message,
        'account_id': account_id,
        'account_email': account.get('email', ''),
        'forward_last_checked_at': cursor_value,
        'lookback_minutes': effective_lookback,
        'triggered': triggered,
    })


@app.route('/api/settings/test-forward-channel', methods=['POST'])
@login_required
def api_test_forward_channel():
    data = request.json or {}
    channel = str(data.get('channel', '') or '').strip().lower()
    config = data.get('config', {}) or {}

    subject = f'[测试消息] 转发链路检测 {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
    body_text = (
        '这是一条由系统主动发送的测试消息。\n'
        '如果你收到了这条消息，说明当前转发链路配置可用。'
    )
    body_html = (
        '<p>这是一条由系统主动发送的测试消息。</p>'
        '<p>如果你收到了这条消息，说明当前转发链路配置可用。</p>'
    )
    telegram_text = f'{subject}\n\n这是一条由系统主动发送的测试消息。\n如果你收到了这条消息，说明当前转发链路配置可用。'

    try:
        if channel == 'smtp':
            smtp_config = config.get('smtp', {}) if isinstance(config, dict) else {}
            if not send_forward_email_with_config(smtp_config, subject, body_text, body_html):
                return jsonify({'success': False, 'error': 'SMTP 测试发送失败，请检查当前表单配置'})
            return jsonify({'success': True, 'message': 'SMTP 测试消息已发送，请检查收件箱'})

        if channel == 'telegram':
            telegram_config = config.get('telegram', {}) if isinstance(config, dict) else {}
            if not send_forward_telegram_with_config(telegram_config, telegram_text):
                return jsonify({'success': False, 'error': 'Telegram 测试发送失败，请检查当前表单配置'})
            return jsonify({'success': True, 'message': 'Telegram 测试消息已发送，请检查目标会话'})

        return jsonify({'success': False, 'error': '未知转发渠道'})
    except Exception as exc:
        return jsonify({'success': False, 'error': f'测试失败: {str(exc)}'})


def init_scheduler():
    """初始化定时任务调度器"""
    global scheduler_instance

    with scheduler_lock:
        if scheduler_instance is not None:
            return scheduler_instance

        try:
            from apscheduler.schedulers.background import BackgroundScheduler
            from apscheduler.triggers.cron import CronTrigger
            import atexit

            scheduler = BackgroundScheduler()

            with app.app_context():
                enable_scheduled = get_setting('enable_scheduled_refresh', 'true').lower() == 'true'

                if not enable_scheduled:
                    print("[OK] 定时刷新已禁用")
                    return None

                use_cron = get_setting('use_cron_schedule', 'false').lower() == 'true'

                if use_cron:
                    cron_expr = get_setting('refresh_cron', '0 2 * * *')
                    try:
                        from croniter import croniter
                        from datetime import datetime
                        croniter(cron_expr, datetime.now())

                        parts = cron_expr.split()
                        if len(parts) == 5:
                            minute, hour, day, month, day_of_week = parts
                            trigger = CronTrigger(
                                minute=minute,
                                hour=hour,
                                day=day,
                                month=month,
                                day_of_week=day_of_week
                            )
                            scheduler.add_job(
                                func=scheduled_refresh_task,
                                trigger=trigger,
                                id='token_refresh',
                                name='Token 定时刷新',
                                replace_existing=True
                            )
                            forward_interval = max(1, min(60, int(get_setting('forward_check_interval_minutes', '5') or '5')))
                            scheduler.add_job(
                                func=process_forwarding_job,
                                trigger=CronTrigger(minute=f'*/{forward_interval}'),
                                id='forward_mail',
                                name='邮件转发轮询',
                                replace_existing=True
                            )
                            scheduler.start()
                            scheduler_instance = scheduler
                            print(f"[OK] 定时任务已启动：Cron 表达式 '{cron_expr}'")
                            atexit.register(lambda: scheduler.shutdown())
                            return scheduler_instance
                        else:
                            print("[WARN] Cron 表达式格式错误，回退到默认配置")
                    except Exception as e:
                        print(f"[WARN] Cron 表达式解析失败: {str(e)}，回退到默认配置")

                refresh_interval_days = int(get_setting('refresh_interval_days', '30'))
                scheduler.add_job(
                    func=scheduled_refresh_task,
                    trigger=CronTrigger(hour=2, minute=0),
                    id='token_refresh',
                    name='Token 定时刷新',
                    replace_existing=True
                )

                forward_interval = max(1, min(60, int(get_setting('forward_check_interval_minutes', '5') or '5')))
                scheduler.add_job(
                    func=process_forwarding_job,
                    trigger=CronTrigger(minute=f'*/{forward_interval}'),
                    id='forward_mail',
                    name='邮件转发轮询',
                    replace_existing=True
                )
                scheduler.start()
                scheduler_instance = scheduler
                print(f"[OK] 定时任务已启动：每天凌晨 2:00 检查刷新（周期：{refresh_interval_days} 天）")

            atexit.register(lambda: scheduler.shutdown())

            return scheduler_instance
        except ImportError:
            print("[WARN] APScheduler 未安装，定时任务功能不可用")
            print("  安装命令：pip install APScheduler>=3.10.0")
            return None
        except Exception as e:
            print(f"[WARN] 定时任务初始化失败：{str(e)}")
            return None


def shutdown_scheduler():
    """关闭定时任务调度器。"""
    global scheduler_instance

    with scheduler_lock:
        if scheduler_instance is None:
            return
        try:
            scheduler_instance.shutdown(wait=False)
        except Exception:
            pass
        scheduler_instance = None


def ensure_scheduler_started():
    """确保调度器已启动（兼容 gunicorn / docker compose）"""
    if os.getenv('WERKZEUG_RUN_MAIN') == 'false':
        return None
    return init_scheduler()


def scheduled_refresh_task():
    """定时刷新任务（由调度器调用）"""
    from datetime import datetime, timedelta

    try:
        with app.app_context():
            enable_scheduled = get_setting('enable_scheduled_refresh', 'true').lower() == 'true'

            if not enable_scheduled:
                print(f"[定时任务] 定时刷新已禁用，跳过执行")
                return

            use_cron = get_setting('use_cron_schedule', 'false').lower() == 'true'

            if use_cron:
                print(f"[定时任务] 使用 Cron 调度，直接执行刷新...")
                trigger_refresh_internal()
                print(f"[定时任务] Token 刷新完成")
                return

            refresh_interval_days = int(get_setting('refresh_interval_days', '30'))

        conn = connect_database()
        cursor = conn.execute('''
            SELECT MAX(created_at) as last_refresh
            FROM account_refresh_logs
            WHERE refresh_type = 'scheduled'
        ''')
        row = cursor.fetchone()
        conn.close()

        last_refresh = row['last_refresh'] if row and row['last_refresh'] else None

        if last_refresh:
            last_refresh_time = datetime.fromisoformat(last_refresh)
            next_refresh_time = last_refresh_time + timedelta(days=refresh_interval_days)
            if datetime.now() < next_refresh_time:
                print(f"[定时任务] 距离上次刷新未满 {refresh_interval_days} 天，跳过本次刷新")
                return

        print(f"[定时任务] 开始执行 Token 刷新...")
        trigger_refresh_internal()
        print(f"[定时任务] Token 刷新完成")

    except Exception as e:
        print(f"[定时任务] 执行失败：{str(e)}")


ensure_scheduler_started()


def trigger_refresh_internal():
    """内部触发刷新（不通过 HTTP）"""
    conn = connect_database()

    try:
        # 获取刷新间隔配置
        cursor_settings = conn.execute("SELECT value FROM settings WHERE key = 'refresh_delay_seconds'")
        delay_row = cursor_settings.fetchone()
        delay_seconds = int(delay_row['value']) if delay_row else 5

        # 清理超过半年的刷新记录
        conn.execute("DELETE FROM account_refresh_logs WHERE created_at < datetime('now', '-6 months')")
        conn.commit()

        accounts = load_refresh_target_accounts(conn)

        total = len(accounts)
        success_count = 0
        failed_count = 0

        for index, account in enumerate(accounts, 1):
            account_id = account['id']
            account_email = account['email']
            client_id = account['client_id']
            encrypted_refresh_token = account['refresh_token']

            try:
                refresh_token = decrypt_data(encrypted_refresh_token) if encrypted_refresh_token else encrypted_refresh_token
            except Exception as e:
                failed_count += 1
                error_msg = f"解密 token 失败: {str(e)}"
                conn.execute('''
                    INSERT INTO account_refresh_logs (account_id, account_email, refresh_type, status, error_message)
                    VALUES (?, ?, ?, ?, ?)
                ''', (account_id, account_email, 'scheduled', 'failed', error_msg))
                update_account_status_by_refresh_result(conn, account_id, False)
                conn.commit()
                continue

            # 获取分组代理设置
            proxy_url = ''
            group_id = account['group_id']
            group_row = None
            if group_id:
                group_cursor = conn.execute(
                    'SELECT proxy_url, fallback_proxy_url_1, fallback_proxy_url_2 FROM groups WHERE id = ?',
                    (group_id,)
                )
                group_row = group_cursor.fetchone()
                if group_row:
                    proxy_url = group_row['proxy_url'] or ''
            fallback_proxy_urls = get_group_proxy_failover_urls(dict(group_row) if group_row else None)

            probe_result = probe_account_permissions(
                account_email, client_id, refresh_token, proxy_url, fallback_proxy_urls,
                account['mail_protocol'], account['oauth_authority']
            )
            success = bool(probe_result.get('success'))
            error_msg = probe_result.get('error', '') if not success else None

            conn.execute('''
                INSERT INTO account_refresh_logs (account_id, account_email, refresh_type, status, error_message)
                VALUES (?, ?, ?, ?, ?)
            ''', (account_id, account_email, 'scheduled', 'success' if success else 'failed', error_msg))
            update_account_status_by_refresh_result(conn, account_id, success)
            save_account_permission_probe(conn, account_id, probe_result)

            if success:
                success_count += 1
            else:
                failed_count += 1

            conn.commit()

            if index < total and delay_seconds > 0:
                time.sleep(delay_seconds)

        print(f"[定时任务] 刷新结果：总计 {total}，成功 {success_count}，失败 {failed_count}")

    finally:
        conn.close()


# ==================== 错误处理 ====================

def api_update_account_v2(account_id):
    data = request.json or {}

    if 'status' in data and len(data) == 1:
        return api_update_account_status(account_id, data['status'])

    email_addr = (data.get('email', '') or '').strip()
    password = data.get('password', '') or ''
    client_id = (data.get('client_id', '') or '').strip()
    refresh_token = (data.get('refresh_token', '') or '').strip()
    account_type = (data.get('account_type', 'outlook') or 'outlook').strip().lower()
    provider = (data.get('provider', 'outlook') or 'outlook').strip().lower()
    mail_protocol = normalize_mail_protocol(data.get('mail_protocol'))
    oauth_authority = normalize_oauth_authority(data.get('oauth_authority'))
    imap_host = (data.get('imap_host', '') or '').strip()
    imap_password = data.get('imap_password', '') or ''
    group_id = data.get('group_id', 1)
    remark = sanitize_input(data.get('remark', ''), max_length=200)
    status = data.get('status', 'active')
    forward_enabled = bool(data.get('forward_enabled', False))
    aliases_provided = 'aliases' in data
    aliases = parse_alias_payload(data.get('aliases', [])) if aliases_provided else []

    try:
        group_id = int(group_id or 1)
    except (TypeError, ValueError):
        group_id = 1

    try:
        imap_port = int(data.get('imap_port', 993) or 993)
    except (TypeError, ValueError):
        return jsonify({'success': False, 'error': 'IMAP 端口无效'})

    provider_meta = get_provider_meta(provider, email_addr)
    is_outlook = account_type == 'outlook' or provider_meta['key'] == 'outlook'

    if is_outlook:
        if not email_addr or not client_id or not refresh_token:
            return jsonify({'success': False, 'error': '邮箱、Client ID 和 Refresh Token 不能为空'})
        account_type = 'outlook'
        provider = 'outlook'
        imap_host = IMAP_SERVER_NEW
        imap_port = IMAP_PORT
        imap_password = ''
    else:
        if not email_addr or not imap_password:
            return jsonify({'success': False, 'error': '邮箱和 IMAP 密码不能为空'})
        account_type = 'imap'
        provider = provider_meta['key']
        client_id = ''
        refresh_token = ''
        password = ''
        if provider == 'custom':
            if not imap_host:
                return jsonify({'success': False, 'error': '自定义 IMAP 必须填写服务器地址'})
        else:
            imap_host = provider_meta.get('imap_host', '')
            imap_port = int(provider_meta.get('imap_port', 993) or 993)

    if aliases_provided:
        _, alias_errors = validate_account_aliases(account_id, email_addr, aliases)
        if alias_errors:
            return jsonify({'success': False, 'error': '；'.join(alias_errors), 'errors': alias_errors})

    if update_account(
        account_id,
        email_addr,
        password,
        client_id,
        refresh_token,
        group_id,
        remark,
        status,
        account_type,
        provider,
        imap_host,
        imap_port,
        imap_password,
        forward_enabled,
        mail_protocol,
        oauth_authority
    ):
        cleaned_aliases = get_account_aliases(account_id)
        if aliases_provided:
            db = get_db()
            alias_success, cleaned_aliases, alias_errors = replace_account_aliases(account_id, email_addr, aliases, db)
            if not alias_success:
                db.rollback()
                return jsonify({'success': False, 'error': '；'.join(alias_errors), 'errors': alias_errors})
            db.commit()
        return jsonify({'success': True, 'message': '账号更新成功', 'aliases': cleaned_aliases})
    return jsonify({'success': False, 'error': '更新失败'})


def api_get_emails_v2(email_addr):
    account = get_account_by_email(email_addr)
    if not account:
        error_payload = build_error_payload(
            "ACCOUNT_NOT_FOUND",
            "账号不存在",
            "NotFoundError",
            404,
            f"email={email_addr}"
        )
        return jsonify({'success': False, 'error': error_payload})

    folder = normalize_folder_name(request.args.get('folder', 'inbox'))
    skip = int(request.args.get('skip', 0))
    top = int(request.args.get('top', 20))
    subject_contains = request.args.get('subject_contains', '').strip().lower()
    from_contains = request.args.get('from_contains', '').strip().lower()
    keyword = request.args.get('keyword', '').strip().lower()
    result = fetch_account_emails(account, folder, skip, top)
    if result.get('success'):
        if subject_contains or from_contains or keyword:
            result['emails'] = [
                item for item in result.get('emails', [])
                if email_matches_filters(account, item, subject_contains, from_contains, keyword)
            ]
        db = get_db()
        db.execute(
            '''
            UPDATE accounts
            SET last_refresh_at = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            ''',
            (account['id'],)
        )
        db.commit()
    return jsonify(result)


def email_matches_filters(account: Dict[str, Any], item: Dict[str, Any],
                          subject_contains: str = '', from_contains: str = '',
                          keyword: str = '') -> bool:
    subject = str(item.get('subject', '') or '')
    sender = str(item.get('from', '') or '')
    preview = str(item.get('body_preview', '') or '')

    if subject_contains and subject_contains not in subject.lower():
        return False
    if from_contains and from_contains not in sender.lower():
        return False
    if not keyword:
        return True

    base_text = '\n'.join([subject, preview]).lower()
    if keyword in base_text:
        return True

    proxy_url = get_account_proxy_url(account)
    fallback_proxy_urls = get_account_proxy_failover_urls(account)
    if account.get('account_type') == 'imap':
        detail_payload = get_email_detail_imap_generic_result(
            account['email'],
            account.get('imap_password', ''),
            account.get('imap_host', ''),
            account.get('imap_port', 993),
            str(item.get('id', '')),
            item.get('folder', 'inbox'),
            account.get('provider', 'custom'),
            proxy_url
        )
        if detail_payload and detail_payload.get('success'):
            body = str((detail_payload.get('email') or {}).get('body', '') or '')
            return keyword in strip_html_content(body).lower()
        return False

    detail = get_email_detail_graph(
        account.get('client_id', ''),
        account.get('refresh_token', ''),
        str(item.get('id', '')),
        proxy_url,
        fallback_proxy_urls,
    )
    if not detail:
        return False
    body = str((detail.get('body') or {}).get('content', '') or '')
    return keyword in strip_html_content(body).lower()


def read_account_messages(account: Dict[str, Any], folder: str = 'inbox', **kwargs) -> Dict[str, Any]:
    normalized_folder = normalize_folder_name(folder)
    skip = int(kwargs.get('skip', 0) or 0)
    top = int(kwargs.get('top', 20) or 20)
    return fetch_account_emails(account, normalized_folder, skip, top)


def read_account_message_detail(account: Dict[str, Any], message_id: str, folder: str = 'inbox') -> Dict[str, Any]:
    folder_name = normalize_folder_name(folder)
    proxy_url = get_account_proxy_url(account)
    fallback_proxy_urls = get_account_proxy_failover_urls(account)
    mail_protocol = normalize_mail_protocol(account.get('mail_protocol'))
    oauth_authority = normalize_oauth_authority(account.get('oauth_authority'))

    if account.get('account_type') == 'imap':
        result = get_email_detail_imap_generic_result(
            account['email'],
            account.get('imap_password', ''),
            account.get('imap_host', ''),
            account.get('imap_port', 993),
            message_id,
            folder_name,
            account.get('provider', 'custom'),
            proxy_url,
        )
        if not result.get('success'):
            return result
        email_detail = result.get('email') or {}
        body = str(email_detail.get('body', '') or '')
        body_type = str(email_detail.get('body_type', 'text') or 'text').lower()
        return {
            'success': True,
            'email': {
                'id': str(email_detail.get('id', '') or message_id),
                'subject': email_detail.get('subject', '无主题'),
                'content': strip_html_content(body),
                'html_content': body if body_type == 'html' else '',
                'from_address': email_detail.get('from', '未知'),
                'timestamp': email_detail.get('date', ''),
                'method': 'IMAP',
            }
        }

    detail = None
    if mail_protocol in ('auto', 'graph'):
        detail = get_email_detail_graph(
            account.get('client_id', ''),
            account.get('refresh_token', ''),
            str(message_id),
            proxy_url,
            fallback_proxy_urls,
            oauth_authority,
        )
    if not detail and mail_protocol in ('auto', 'imap_oauth'):
        imap_detail = get_email_detail_imap(
            account.get('email', ''),
            account.get('client_id', ''),
            account.get('refresh_token', ''),
            str(message_id),
            folder_name,
            proxy_url,
            fallback_proxy_urls,
            oauth_authority,
        )
        if imap_detail:
            body = str(imap_detail.get('body', '') or '')
            return {
                'success': True,
                'email': {
                    'id': str(imap_detail.get('id', '') or message_id),
                    'subject': imap_detail.get('subject', '无主题'),
                    'content': strip_html_content(body),
                    'html_content': '',
                    'from_address': imap_detail.get('from', '未知'),
                    'timestamp': imap_detail.get('date', ''),
                    'method': 'IMAP (OAuth)',
                }
            }
    if not detail:
        return {'success': False, 'error': '获取邮件详情失败'}

    html_content = str((detail.get('body') or {}).get('content', '') or '')
    return {
        'success': True,
        'email': {
            'id': str(detail.get('id', '') or message_id),
            'subject': detail.get('subject', '无主题'),
            'content': strip_html_content(html_content),
            'html_content': html_content,
            'from_address': (detail.get('from') or {}).get('emailAddress', {}).get('address', '未知'),
            'timestamp': detail.get('receivedDateTime', ''),
            'method': 'Graph API',
        }
    }


def extract_verification_code(content: str) -> str:
    match = re.search(r'(?<!\d)(\d{6})(?!\d)', str(content or ''))
    return match.group(1) if match else ''


def account_has_imap_otp_credentials(account: Dict[str, Any]) -> bool:
    return (str(account.get('account_type', '') or '').strip().lower() == 'imap')


def can_attempt_imap_otp(account: Dict[str, Any]) -> bool:
    account_type = str(account.get('account_type', '') or '').strip().lower()
    if account_type == 'imap':
        return True
    imap_password = str(account.get('imap_password', '') or '').strip()
    imap_host = str(account.get('imap_host', '') or '').strip()
    oauth_client_id = str(account.get('client_id', '') or '').strip()
    oauth_refresh_token = str(account.get('refresh_token', '') or '').strip()
    return bool((imap_password and imap_host) or (oauth_client_id and oauth_refresh_token))


def fetch_verification_candidates_imap(account: Dict[str, Any], folder: str, top: int = 20) -> Dict[str, Any]:
    folder_name = normalize_folder_name(folder)
    proxy_url = get_account_proxy_url(account)
    fallback_proxy_urls = get_account_proxy_failover_urls(account)
    provider = str(account.get('provider', '') or '').strip().lower()
    oauth_client_id = str(account.get('client_id', '') or '').strip()
    oauth_refresh_token = str(account.get('refresh_token', '') or '').strip()
    oauth_error = None

    if provider == 'outlook' and oauth_client_id and oauth_refresh_token:
        oauth_result = get_emails_imap_with_server(
            account['email'],
            oauth_client_id,
            oauth_refresh_token,
            folder_name,
            0,
            top,
            IMAP_SERVER_NEW,
            proxy_url,
            fallback_proxy_urls,
        )
        if oauth_result.get('success'):
            formatted = format_email_items(oauth_result.get('emails', []), folder_name)
            return {
                'success': True,
                'emails': [dict(item, folder=folder_name) for item in formatted],
                'method': 'IMAP (OAuth)',
            }
        oauth_error = oauth_result.get('error')

    result = get_emails_imap_generic(
        account['email'],
        account.get('imap_password', ''),
        account.get('imap_host', ''),
        account.get('imap_port', 993),
        folder_name,
        account.get('provider', 'custom'),
        0,
        top,
        proxy_url,
    )
    if not result.get('success'):
        return {
            'success': False,
            'error': oauth_error or result.get('error') or '获取邮件失败',
        }

    return {
        'success': True,
        'emails': [dict(item, folder=folder_name) for item in result.get('emails', [])],
        'method': result.get('method', 'IMAP (Generic)'),
    }


def fetch_verification_detail_imap(account: Dict[str, Any], message_id: str, folder: str = 'inbox') -> Dict[str, Any]:
    folder_name = normalize_folder_name(folder)
    proxy_url = get_account_proxy_url(account)
    fallback_proxy_urls = get_account_proxy_failover_urls(account)
    provider = str(account.get('provider', '') or '').strip().lower()
    oauth_client_id = str(account.get('client_id', '') or '').strip()
    oauth_refresh_token = str(account.get('refresh_token', '') or '').strip()

    if provider == 'outlook' and oauth_client_id and oauth_refresh_token:
        oauth_detail = get_email_detail_imap(
            account['email'],
            oauth_client_id,
            oauth_refresh_token,
            message_id,
            folder_name,
            proxy_url,
            fallback_proxy_urls,
        )
        if oauth_detail:
            body = str(oauth_detail.get('body', '') or '')
            return {
                'success': True,
                'email': {
                    'id': str(oauth_detail.get('id', '') or message_id),
                    'subject': oauth_detail.get('subject', '无主题'),
                    'content': strip_html_content(body),
                    'html_content': '',
                    'from_address': oauth_detail.get('from', '未知'),
                    'timestamp': oauth_detail.get('date', ''),
                    'method': 'IMAP (OAuth)',
                }
            }

    result = get_email_detail_imap_generic_result(
        account['email'],
        account.get('imap_password', ''),
        account.get('imap_host', ''),
        account.get('imap_port', 993),
        message_id,
        folder_name,
        account.get('provider', 'custom'),
        proxy_url,
    )
    if not result.get('success'):
        return result

    email_detail = result.get('email') or {}
    body = str(email_detail.get('body', '') or '')
    body_type = str(email_detail.get('body_type', 'text') or 'text').lower()
    return {
        'success': True,
        'email': {
            'id': str(email_detail.get('id', '') or message_id),
            'subject': email_detail.get('subject', '无主题'),
            'content': strip_html_content(body),
            'html_content': body if body_type == 'html' else '',
            'from_address': email_detail.get('from', '未知'),
            'timestamp': email_detail.get('date', ''),
            'method': 'IMAP',
        }
    }


def find_latest_verification_code(account: Dict[str, Any], since_minutes: Optional[int] = None) -> Dict[str, Any]:
    prefer_imap_path = account_has_imap_otp_credentials(account)
    allow_imap_fallback = can_attempt_imap_otp(account)
    candidates = []
    errors = []
    for folder in ('inbox', 'junkemail'):
        primary_source = 'imap' if prefer_imap_path else 'graph'
        fallback_source = 'graph' if primary_source == 'imap' else 'imap'

        if primary_source == 'imap':
            result = fetch_verification_candidates_imap(account, folder, top=1)
        else:
            fetch_kwargs = {'skip': 0, 'top': 1}
            if since_minutes is not None:
                fetch_kwargs['since_minutes'] = since_minutes
            result = read_account_messages(account, folder=folder, **fetch_kwargs)

        if (not result.get('success')) and fallback_source == 'imap' and allow_imap_fallback:
            result = fetch_verification_candidates_imap(account, folder, top=1)
            primary_source = 'imap'
        elif (not result.get('success')) and fallback_source == 'graph':
            fetch_kwargs = {'skip': 0, 'top': 1}
            if since_minutes is not None:
                fetch_kwargs['since_minutes'] = since_minutes
            result = read_account_messages(account, folder=folder, **fetch_kwargs)
            primary_source = 'graph'

        if not result.get('success'):
            errors.append(result.get('error') or '获取邮件失败')
            continue

        folder_emails = result.get('emails', []) or []
        if folder_emails:
            candidate = dict(folder_emails[0] or {})
            candidate['folder'] = folder
            candidate['_otp_source'] = primary_source
            candidates.append(candidate)

    if not candidates:
        if errors:
            return {'success': False, 'error': errors[0]}
        return {'success': False, 'error': '未找到邮件'}

    candidates.sort(
        key=lambda item: parse_email_datetime(str(item.get('date', '') or '')) or datetime.min,
        reverse=True,
    )

    selected = candidates[0]
    selected_source = str(selected.get('_otp_source', '') or '').strip().lower()
    if selected_source == 'imap':
        detail_result = fetch_verification_detail_imap(
            account,
            str(selected.get('id', '')),
            folder=selected.get('folder', 'inbox'),
        )
        if not detail_result.get('success'):
            detail_result = read_account_message_detail(
                account,
                str(selected.get('id', '')),
                folder=selected.get('folder', 'inbox'),
            )
    else:
        detail_result = read_account_message_detail(
            account,
            str(selected.get('id', '')),
            folder=selected.get('folder', 'inbox'),
        )
        if (not detail_result.get('success')) and allow_imap_fallback:
            detail_result = fetch_verification_detail_imap(
                account,
                str(selected.get('id', '')),
                folder=selected.get('folder', 'inbox'),
            )

    if not detail_result.get('success'):
        return {'success': False, 'error': detail_result.get('error') or '获取邮件详情失败'}

    email_detail = detail_result.get('email') or {}
    joined_text = '\n'.join([
        str(selected.get('subject', '') or ''),
        str(email_detail.get('subject', '') or ''),
        str(email_detail.get('content', '') or ''),
        strip_html_content(str(email_detail.get('html_content', '') or '')),
    ])
    code = extract_verification_code(joined_text)

    if not code:
        return {'success': False, 'error': '未提取到验证码'}

    return {
        'success': True,
        'data': {
            'code': code,
            'selected_folder': selected.get('folder', 'inbox'),
            'selected_message_id': str(selected.get('id', '') or ''),
            'selected_date': selected.get('date', ''),
            'candidates': [
                {
                    'id': str(item.get('id', '') or ''),
                    'folder': item.get('folder', 'inbox'),
                    'date': item.get('date', ''),
                    'subject': item.get('subject', ''),
                }
                for item in candidates
            ],
        }
    }


app.view_functions['api_update_account'] = api_update_account_v2
app.view_functions['api_get_emails'] = api_get_emails_v2


@app.errorhandler(400)
def bad_request(error):
    """处理400错误"""
    print(f"400 Bad Request: {error}")
    return jsonify({'success': False, 'error': '请求格式错误'}), 400


@app.errorhandler(Exception)
def handle_exception(error):
    """处理未捕获的异常"""
    print(f"Unhandled exception: {error}")
    import traceback
    traceback.print_exc()
    return jsonify({'success': False, 'error': str(error)}), 500
