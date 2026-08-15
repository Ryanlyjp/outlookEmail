const $ = selector => document.querySelector(selector);
const $$ = selector => Array.from(document.querySelectorAll(selector));
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let accounts = [];
let selected = new Set();
let sortState = { key: 'id', direction: 'desc' };
let page = 1;
let pageSize = 50;
let csrfToken = null;

async function initCsrfToken() {
    const response = await fetch('/api/csrf-token');
    const data = await response.json();
    csrfToken = data.csrf_token || null;
}

async function api(path, options = {}) {
    const method = String(options.method || 'GET').toUpperCase();
    if (method !== 'GET' && csrfToken) {
        options.headers = { ...(options.headers || {}), 'X-CSRFToken': csrfToken };
    }
    const response = await fetch(path, options);
    const data = await response.json().catch(() => ({}));
    if (response.status === 401) location.href = '/login';
    if (!response.ok || data.success === false) throw new Error(data.error?.message || data.error || '请求失败');
    return data;
}

function toast(message) {
    const el = $('#toast');
    el.textContent = message;
    el.classList.remove('hidden');
    clearTimeout(el._timer);
    el._timer = setTimeout(() => el.classList.add('hidden'), 2600);
}

function healthLabel(item) {
    return ({all:'Graph + IMAP',graph_only:'仅 Graph',imap_only:'仅 IMAP',token_invalid:'Token 失效',unavailable:'协议不可用'})[item.health_status] || '未测试';
}

function chip(text, status) {
    let cls = '';
    if (status === 'ok' || status === 'all') cls = 'ok';
    else if (status === 'failed' || status === 'token_invalid' || status === 'unavailable') cls = 'fail';
    else if (status) cls = 'warn';
    return `<span class="chip ${cls}">${esc(text)}</span>`;
}

function isAvailable(item) {
    return item.status === 'active' && ['all', 'graph_only', 'imap_only'].includes(item.health_status);
}

function updateStats() {
    const set = (id, value) => { $(id).textContent = value; };
    set('#statTotal', accounts.length);
    set('#statAvailable', accounts.filter(isAvailable).length);
    set('#statAll', accounts.filter(a => a.health_status === 'all').length);
    set('#statGraph', accounts.filter(a => a.health_status === 'graph_only').length);
    set('#statImap', accounts.filter(a => a.health_status === 'imap_only').length);
    set('#statInvalid', accounts.filter(a => a.health_status === 'token_invalid').length);
    set('#statUnavailable', accounts.filter(a => a.health_status === 'unavailable').length);
    set('#statUntested', accounts.filter(a => !a.health_status).length);
}

function fillGroups() {
    const current = $('#groupFilter').value;
    const groups = [...new Set(accounts.map(a => a.group_name || '默认分组'))].sort();
    $('#groupFilter').innerHTML = '<option value="">全部分组</option>' + groups.map(name => `<option value="${esc(name)}">${esc(name)}</option>`).join('');
    $('#groupFilter').value = groups.includes(current) ? current : '';
}

function filteredAccounts() {
    const keyword = $('#searchInput').value.trim().toLowerCase();
    const health = $('#healthFilter').value;
    const domain = $('#domainFilter').value;
    const group = $('#groupFilter').value;
    const rows = accounts.filter(item => {
        if (keyword && ![item.email, item.remark, item.group_name].join(' ').toLowerCase().includes(keyword)) return false;
        const suffix = String(item.email || '').split('@')[1]?.toLowerCase() || '';
        if (domain === '__other__' && ['outlook.com','hotmail.com','live.com'].includes(suffix)) return false;
        if (domain && domain !== '__other__' && suffix !== domain) return false;
        if (group && (item.group_name || '默认分组') !== group) return false;
        if (health === 'available') return isAvailable(item);
        if (health === 'all_access') return item.health_status === 'all';
        if (health === 'untested') return !item.health_status;
        if (health !== 'all') return item.health_status === health;
        return true;
    });
    rows.sort((a, b) => {
        const key = sortState.key;
        const av = key === 'id' ? Number(a[key] || 0) : String(a[key] || '').toLowerCase();
        const bv = key === 'id' ? Number(b[key] || 0) : String(b[key] || '').toLowerCase();
        if (av === bv) return 0;
        const result = av < bv ? -1 : 1;
        return sortState.direction === 'asc' ? result : -result;
    });
    return rows;
}

function render() {
    const filtered = filteredAccounts();
    const pages = Math.max(1, Math.ceil(filtered.length / pageSize));
    page = Math.min(Math.max(page, 1), pages);
    const start = (page - 1) * pageSize;
    const rows = filtered.slice(start, start + pageSize);
    $('#accountTable').innerHTML = rows.length ? rows.map(item => `<tr>
        <td><input class="row-check" type="checkbox" data-id="${item.id}" ${selected.has(item.id) ? 'checked' : ''}></td>
        <td>${item.id}</td><td class="email">${esc(item.email)}</td>
        <td><div class="stack">${chip(healthLabel(item), item.health_status)}<span class="reason" title="${esc(item.health_detail || '')}">${esc(item.health_detail || '-')}</span></div></td>
        <td><div class="stack">${chip(`Graph ${item.graph_status || '未测'}`, item.graph_status)}${chip(`IMAP ${item.imap_status || '未测'}`, item.imap_status)}</div></td>
        <td>${esc(item.mail_protocol || 'auto')} / ${esc(item.oauth_authority || 'common')}</td><td>${esc(item.group_name || '默认分组')}</td>
        <td>${esc(item.last_protocol_test_at || '-')}</td>
        <td><div class="actions"><button data-action="test" data-id="${item.id}">测试</button><button data-action="copy2" data-id="${item.id}">复制账密</button><button data-action="copy4" data-id="${item.id}">复制全部</button></div></td>
    </tr>`).join('') : '<tr><td colspan="9" class="empty">没有匹配账号</td></tr>';
    $('#selectedCount').textContent = `已选 ${selected.size}`;
    $('#tableSummary').textContent = `筛选结果 ${filtered.length} 个 · 本页 ${rows.length} 个 · 总库存 ${accounts.length}`;
    $('#pageInfo').textContent = `${page} / ${pages}`;
    $('#prevPage').disabled = page <= 1;
    $('#nextPage').disabled = page >= pages;
    $('#selectPage').checked = rows.length > 0 && rows.every(a => selected.has(a.id));
}

async function loadAccounts() {
    const data = await api('/api/accounts');
    accounts = data.accounts || [];
    selected.forEach(id => { if (!accounts.some(a => a.id === id)) selected.delete(id); });
    updateStats(); fillGroups(); render();
}

async function copyAccount(id, full) {
    const data = await api(`/api/accounts/${id}`);
    const a = data.account;
    const value = full ? `${a.email || ''}----${a.password || ''}----${a.client_id || ''}----${a.refresh_token || ''}` : `${a.email || ''}----${a.password || ''}`;
    if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(value);
    else {
        const textarea = document.createElement('textarea');
        textarea.value = value; textarea.style.position = 'fixed'; textarea.style.opacity = '0';
        document.body.appendChild(textarea); textarea.select(); document.execCommand('copy'); textarea.remove();
    }
    toast(full ? '完整账号已复制' : '账号密码已复制');
}

async function testOne(id) {
    const data = await api(`/api/accounts/${id}/refresh`, { method: 'POST' });
    toast(data.message || '权限测试完成');
    await loadAccounts();
}

async function testSelected() {
    if (!selected.size) return toast('请先选择账号');
    const data = await api('/api/accounts/refresh-selected', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({account_ids:[...selected]}) });
    toast(`测试完成：成功 ${data.success_count}，失败 ${data.failed_count}，跳过 ${data.skipped_count}`);
    await loadAccounts();
}

function showJob(title) {
    $('#jobTitle').textContent = title; $('#jobState').textContent = ''; $('#jobText').textContent = '正在初始化…'; $('#jobBar').style.width = '0%'; $('#jobCard').classList.remove('hidden');
}

function testAll() {
    if (!confirm('确定测试全部 Outlook 账号的 Graph/IMAP 权限吗？')) return;
    showJob('全量权限测试');
    $('#testAllBtn').disabled = true;
    const source = new EventSource('/api/accounts/trigger-scheduled-refresh?force=true');
    source.onmessage = async event => {
        const data = JSON.parse(event.data);
        if (data.type === 'start') $('#jobState').textContent = `共 ${data.total} 个账号`;
        if (data.type === 'progress') {
            $('#jobBar').style.width = `${Math.round(data.current / Math.max(data.total, 1) * 100)}%`;
            $('#jobText').textContent = `${data.current}/${data.total} · ${data.email} · 成功 ${data.success_count} · 失败 ${data.failed_count}`;
        }
        if (data.type === 'complete') {
            source.close(); $('#testAllBtn').disabled = false; $('#jobBar').style.width = '100%'; $('#jobText').textContent = `完成：成功 ${data.success_count}，失败 ${data.failed_count}`; await loadAccounts();
        }
    };
    source.onerror = () => { source.close(); $('#testAllBtn').disabled = false; $('#jobText').textContent = '任务连接中断'; };
}

async function retryFailed() {
    showJob('重试失败账号'); $('#retryFailedBtn').disabled = true;
    try { const data = await api('/api/accounts/refresh-failed', {method:'POST'}); $('#jobBar').style.width = '100%'; $('#jobText').textContent = `完成：成功 ${data.success_count || 0}，失败 ${data.failed_count || 0}`; await loadAccounts(); }
    catch (error) { $('#jobText').textContent = error.message; }
    finally { $('#retryFailedBtn').disabled = false; }
}

function openDetails(title, html) { $('#detailTitle').textContent = title; $('#detailContent').innerHTML = html; $('#detailPanel').classList.remove('hidden'); }
async function showFailed() {
    const data = await api('/api/accounts/refresh-logs/failed');
    openDetails('失败账号', data.logs?.length ? data.logs.map(log => `<div class="stack"><strong>${esc(log.account_email)}</strong><span class="reason">${esc(log.error_message || '未知错误')} · ${esc(log.created_at || '')}</span></div><hr>`).join('') : '<div class="empty">暂无失败账号</div>');
}
async function showHistory() {
    const data = await api('/api/accounts/refresh-logs?limit=1000');
    openDetails('测试历史', data.logs?.length ? data.logs.map(log => `<div><strong>${esc(log.account_email)}</strong> ${chip(log.status === 'success' ? '成功' : '失败', log.status === 'success' ? 'ok' : 'failed')}<div class="reason">${esc(log.created_at || '')}${log.error_message ? ` · ${esc(log.error_message)}` : ''}</div></div><hr>`).join('') : '<div class="empty">暂无测试历史</div>');
}

document.addEventListener('DOMContentLoaded', () => {
    initCsrfToken().then(loadAccounts).catch(error => toast(error.message));
    ['searchInput','healthFilter','domainFilter','groupFilter'].forEach(id => $("#"+id).addEventListener(id === 'searchInput' ? 'input' : 'change', () => { page = 1; render(); }));
    $$('#poolStats button').forEach(button => button.addEventListener('click', () => { $('#healthFilter').value = button.dataset.filter; page = 1; render(); }));
    $$('th[data-sort]').forEach(th => th.addEventListener('click', () => { const key=th.dataset.sort; sortState={key,direction:sortState.key===key&&sortState.direction==='asc'?'desc':'asc'}; render(); }));
    $('#reloadBtn').onclick = () => loadAccounts().catch(error => toast(error.message));
    $('#selectFilteredBtn').onclick = () => { filteredAccounts().forEach(a => selected.add(a.id)); render(); };
    $('#clearSelectedBtn').onclick = () => { selected.clear(); render(); };
    $('#testSelectedBtn').onclick = () => testSelected().catch(error => toast(error.message));
    $('#testAllBtn').onclick = testAll; $('#retryFailedBtn').onclick = retryFailed; $('#showFailedBtn').onclick = () => showFailed().catch(error => toast(error.message)); $('#showHistoryBtn').onclick = () => showHistory().catch(error => toast(error.message));
    $('#closeDetailBtn').onclick = () => $('#detailPanel').classList.add('hidden');
    $('#pageSize').onchange = e => { pageSize=Number(e.target.value); page=1; render(); }; $('#prevPage').onclick=()=>{page--;render();}; $('#nextPage').onclick=()=>{page++;render();};
    $('#selectPage').onchange = e => { const rows=filteredAccounts().slice((page-1)*pageSize,page*pageSize); rows.forEach(a=>e.target.checked?selected.add(a.id):selected.delete(a.id)); render(); };
    $('#accountTable').addEventListener('change', e => { if(!e.target.matches('.row-check'))return; const id=Number(e.target.dataset.id); e.target.checked?selected.add(id):selected.delete(id); render(); });
    $('#accountTable').addEventListener('click', e => { const b=e.target.closest('[data-action]'); if(!b)return; const id=Number(b.dataset.id); const run=b.dataset.action==='test'?testOne(id):copyAccount(id,b.dataset.action==='copy4'); run.catch(error=>toast(error.message)); });
});
