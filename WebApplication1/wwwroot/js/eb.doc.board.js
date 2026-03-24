// 2026.03.23 Changed: 협조 문서함 메인탭 서브탭 URL 상태 목록조회 배지갱신을 eb.doc.board.js 전체 흐름에 추가하고 협조 탭 클릭 불가 및 배지 미표시 문제를 수정
// 2026.02.05 Changed: 일괄승인 토스트 수량을 서버 응답 totals approved 기반으로만 표시하도록 수정
// 2026.02.05 Changed: 보드 일괄승인 토스트를 스크린샷과 동일한 상단 중앙 녹색 바 스타일로 통일하고 하단 토스트 로직을 제거

(function () {
    'use strict';

    function $(sel, root) {
        try { return (root || document).querySelector(sel); } catch { return null; }
    }

    function safeParseI18n() {
        try {
            const el = document.getElementById('docBoardI18n');
            if (!el) return {};
            const raw = (el.textContent || '').trim();
            if (!raw) return {};
            const obj = JSON.parse(raw);
            return (obj && typeof obj === 'object') ? obj : {};
        } catch { return {}; }
    }

    function toBool(v) {
        if (v === true) return true;
        if (v === false) return false;
        if (v == null) return false;
        if (typeof v === 'number') return v !== 0;
        if (typeof v === 'string') {
            const s = v.trim().toLowerCase();
            if (s === 'true' || s === '1' || s === 'y' || s === 'yes') return true;
            if (s === 'false' || s === '0' || s === 'n' || s === 'no' || s === '') return false;
        }
        return !!v;
    }

    function pick(obj, keys) {
        if (!obj) return undefined;
        for (var i = 0; i < keys.length; i++) {
            var k = keys[i];
            if (k in obj) return obj[k];
        }
        return undefined;
    }

    function asInt(v) {
        var n = parseInt(String(v ?? '0'), 10);
        return Number.isFinite(n) ? n : 0;
    }

    function fmt0(template, n) {
        try { return String(template || '').replace('{0}', String(n)); } catch { return String(n); }
    }

    function ensureTopToastHost() {
        var host = document.getElementById('ebTopToastHost');
        if (!host) {
            host = document.createElement('div');
            host.id = 'ebTopToastHost';
            host.setAttribute('aria-live', 'polite');
            host.setAttribute('aria-atomic', 'true');
            document.body.appendChild(host);
        }
        return host;
    }

    function clearTopToasts() {
        try {
            var host = ensureTopToastHost();
            host.innerHTML = '';
        } catch { }
    }

    function showTopBarToast(message, kind) {
        try {
            if (!message) return;

            clearTopToasts();

            var host = ensureTopToastHost();

            var box = document.createElement('div');
            box.className = 'eb-topbar-toast';

            if (kind === 'danger') box.classList.add('eb-topbar-toast--danger');
            else box.classList.add('eb-topbar-toast--success');

            var msg = document.createElement('div');
            msg.className = 'eb-topbar-toast__msg';
            msg.textContent = String(message);

            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'eb-topbar-toast__close';
            btn.setAttribute('aria-label', 'Close');
            btn.innerHTML =
                '<svg viewBox="0 0 16 16" aria-hidden="true">' +
                '<path fill="currentColor" d="M3.72 3.72a.75.75 0 0 1 1.06 0L8 6.94l3.22-3.22a.75.75 0 1 1 1.06 1.06L9.06 8l3.22 3.22a.75.75 0 1 1-1.06 1.06L8 9.06l-3.22 3.22a.75.75 0 1 1-1.06-1.06L6.94 8 3.72 4.78a.75.75 0 0 1 0-1.06z"/>' +
                '</svg>';

            btn.addEventListener('click', function () {
                try { box.remove(); } catch { }
            });

            box.appendChild(msg);
            box.appendChild(btn);
            host.appendChild(box);

            setTimeout(function () {
                try { box.remove(); } catch { }
            }, 3500);
        } catch { }
    }

    function createBoard(root) {
        if (!root || !root.dataset) {
            console.error('[DocBoard] root or dataset missing');
            return;
        }

        var apiList = String(root.dataset.apiList || '');
        var apiBadges = String(root.dataset.apiBadges || '');
        var detailUrl = String(root.dataset.detailUrl || '');

        if (!apiList) {
            console.error('[DocBoard] data-api-list is empty. root:', root);
            return;
        }

        var body =
            document.getElementById('docListBody')
            || $('tbody#docListBody', root)
            || $('tbody[data-doc-list-body="1"]', root)
            || $('table tbody', root);

        var paging =
            document.getElementById('docPaging')
            || $('#docPaging', root)
            || $('[data-doc-paging="1"]', root);

        if (!body) {
            console.error('[DocBoard] list tbody not found. id=docListBody or fallback tbody missing. root:', root);
            return;
        }
        if (!paging) {
            console.warn('[DocBoard] paging element not found. pagination UI may be missing.');
            paging = document.createElement('div');
        }

        var tabs = Array.from(root.querySelectorAll('.doc-tab'));

        var createdSubWrap = document.getElementById('createdSubtabs');
        var approvalSubWrap = document.getElementById('approvalSubtabs');
        var cooperationSubWrap = document.getElementById('cooperationSubtabs');
        var sharedSubWrap = document.getElementById('sharedSubtabs');

        var createdSubTabs = createdSubWrap ? Array.from(createdSubWrap.querySelectorAll('.doc-subtab')) : [];
        var approvalSubTabs = approvalSubWrap ? Array.from(approvalSubWrap.querySelectorAll('.doc-subtab')) : [];
        var cooperationSubTabs = cooperationSubWrap ? Array.from(cooperationSubWrap.querySelectorAll('.doc-subtab')) : [];
        var sharedSubTabs = sharedSubWrap ? Array.from(sharedSubWrap.querySelectorAll('.doc-subtab')) : [];

        var filterTitle = document.getElementById('filterTitle') || { value: 'all', querySelectorAll: function () { return []; } };
        var filterSort = document.getElementById('filterSort') || { value: 'created_desc' };
        var filterQuery = document.getElementById('filterQuery') || { value: '' };
        var btnApply = document.getElementById('filterApply');

        var noHeaderWrap = document.getElementById('docNoHeaderWrap');
        var pagingTools = document.getElementById('docPagingTools');
        var chkAllBottom = document.getElementById('chkAllBottom');
        var btnApproveBulk = document.getElementById('btnApproveBulk');

        function findApproveButtonFallback() {
            var b = document.getElementById('btnApproveBulk');
            if (b) return b;

            if (pagingTools) {
                var btns = Array.from(pagingTools.querySelectorAll('button'));
                for (var i = 0; i < btns.length; i++) {
                    var t = (btns[i].textContent || '').trim();
                    if (t === '승인') return btns[i];
                }
            }

            var hook = document.querySelector('button[data-bulk-approve="1"]');
            if (hook) return hook;

            return null;
        }

        var chkAllTop = null;
        var selectedDocIds = new Set();

        var clamp = function (v, min, max) { return Math.max(min, Math.min(max, v)); };
        var allowedTabs = ['created', 'approval', 'cooperation', 'shared'];
        var STORAGE_KEY = 'docBoardStateByTab';
        var SNAPSHOT_PREFIX = 'docBoardSnapshot:';

        var i18n = safeParseI18n();
        var STATUS_LABELS = i18n.STATUS_LABELS || {};
        var READ_LABEL = String(i18n.READ_LABEL || 'Viewed');
        var UNREAD_LABEL = String(i18n.UNREAD_LABEL || 'Unviewed');
        var LIST_LOADING = String(i18n.LIST_LOADING || 'Loading...');
        var LIST_EMPTY = String(i18n.LIST_EMPTY || 'Empty');
        var PAGINATION_PREV = String(i18n.PAGINATION_PREV || 'Prev');
        var PAGINATION_NEXT = String(i18n.PAGINATION_NEXT || 'Next');

        var BULK_OK_FMT = String(i18n.DOC_Toast_BulkApproved || '');
        var BULK_FAIL = String(i18n.BULK_FAIL || '');

        function setBadgeDom(id, val) {
            var el = document.getElementById(id);
            if (!el) return;
            var n = asInt(val);
            el.textContent = String(n);
            el.hidden = !(n > 0);
        }

        function getColSpan() {
            return isApprovalOngoing() ? 7 : 6;
        }

        function getCsrfToken() {
            try {
                var el = document.querySelector('input[name="__RequestVerificationToken"]');
                if (el && el.value) return String(el.value);
            } catch { }
            try {
                var m = document.querySelector('meta[name="RequestVerificationToken"]');
                if (m && m.content) return String(m.content);
            } catch { }
            return '';
        }

        async function doBulkApprove(DocIds) {
            var token = getCsrfToken();
            var payload = { DocIds: DocIds || [] };

            var res = await fetch('/Doc/BulkApprove', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'RequestVerificationToken': token
                },
                credentials: 'same-origin',
                body: JSON.stringify(payload)
            });

            var j = await res.json().catch(function () { return null; });
            if (!res.ok || !j || j.ok !== true) {
                var err = new Error('bulk approve failed');
                err.payload = j;
                err.status = res.status;
                throw err;
            }
            return j;
        }

        function getStatusBase(raw) {
            if (!raw) return 'Pending';
            var s = String(raw).trim().toUpperCase();
            if (s.indexOf('PENDING') === 0) return 'Pending';
            if (s === 'APPROVE' || s.indexOf('APPROVED') === 0) return 'Approved';
            if (s === 'REJECT' || s.indexOf('REJECTED') === 0) return 'Rejected';
            if (s === 'HOLD' || s.indexOf('ONHOLD') === 0 || s.indexOf('ON HOLD') === 0) return 'OnHold';
            if (s === 'RECALL' || s.indexOf('RECALLED') === 0) return 'Recalled';
            return raw;
        }

        function formatStatus(raw, done, total) {
            var base = getStatusBase(raw);
            var label = (STATUS_LABELS && STATUS_LABELS[base]) ? String(STATUS_LABELS[base]) : String(raw || '');
            if (typeof total === 'number' && total > 0) {
                var d = (typeof done === 'number' && done >= 0) ? done : 0;
                return label + ' (' + d + '/' + total + ')';
            }
            return label;
        }

        function readJson(key, fallback) {
            try {
                const s = window.sessionStorage.getItem(key);
                if (!s) return fallback;
                const o = JSON.parse(s);
                return (o && typeof o === 'object') ? o : fallback;
            } catch { return fallback; }
        }

        function writeJson(key, obj) {
            try { window.sessionStorage.setItem(key, JSON.stringify(obj || {})); } catch { }
        }

        function readStateMap() {
            try {
                var s = window.sessionStorage.getItem(STORAGE_KEY);
                if (!s) return {};
                var obj = JSON.parse(s);
                return (obj && typeof obj === 'object') ? obj : {};
            } catch { return {}; }
        }

        function writeStateMap(map) {
            try { window.sessionStorage.setItem(STORAGE_KEY, JSON.stringify(map || {})); } catch { }
        }

        function saveStateToSession() {
            var map = readStateMap();
            map[state.tab] = {
                tab: state.tab,
                page: state.page,
                pageSize: state.pageSize,
                titleFilter: state.titleFilter,
                sort: state.sort,
                q: state.q,
                createdSub: state.createdSub,
                approvalSub: state.approvalSub,
                cooperationSub: state.cooperationSub,
                sharedSub: state.sharedSub
            };
            writeStateMap(map);
        }

        function setUrlFromState(replaceOnly) {
            try {
                var qs = new URLSearchParams(window.location.search || '');
                qs.set('tab', state.tab);
                qs.set('page', String(state.page || 1));
                qs.set('pageSize', String(state.pageSize || 20));
                qs.set('titleFilter', String(state.titleFilter || 'all'));
                qs.set('sort', String(state.sort || 'created_desc'));
                qs.set('q', String(state.q || ''));

                if (state.tab === 'created') qs.set('createdSub', String(state.createdSub || 'ongoing'));
                else qs.delete('createdSub');

                if (state.tab === 'approval') qs.set('approvalSub', String(state.approvalSub || 'ongoing'));
                else qs.delete('approvalSub');

                if (state.tab === 'cooperation') qs.set('cooperationSub', String(state.cooperationSub || 'ongoing'));
                else qs.delete('cooperationSub');

                if (state.tab === 'shared') qs.set('sharedSub', String(state.sharedSub || 'ongoing'));
                else qs.delete('sharedSub');

                var url = window.location.pathname + '?' + qs.toString();
                if (replaceOnly) window.history.replaceState(null, '', url);
                else window.history.pushState(null, '', url);
            } catch { }
        }

        function readStateFromUrl() {
            var qs = new URLSearchParams(window.location.search || '');
            var tab = qs.get('tab');
            var page = qs.get('page');
            var pageSize = qs.get('pageSize');
            var titleFilter = qs.get('titleFilter');
            var sort = qs.get('sort');
            var q = qs.get('q');

            var createdSub = qs.get('createdSub');
            var approvalSub = qs.get('approvalSub');
            var cooperationSub = qs.get('cooperationSub');
            var sharedSub = qs.get('sharedSub');

            var out = {};
            if (tab && allowedTabs.indexOf(tab) >= 0) out.tab = tab;

            var p = parseInt(page || '', 10);
            var ps = parseInt(pageSize || '', 10);
            if (!isNaN(p) && p > 0) out.page = p;
            if (!isNaN(ps) && ps > 0 && ps <= 100) out.pageSize = ps;

            if (titleFilter) out.titleFilter = titleFilter;
            if (sort) out.sort = sort;
            if (q != null) out.q = q;

            if (createdSub) out.createdSub = createdSub;
            if (approvalSub) out.approvalSub = approvalSub;
            if (cooperationSub) out.cooperationSub = cooperationSub;
            if (sharedSub) out.sharedSub = sharedSub;

            return out;
        }

        function readStateForTabFromSession(tab) {
            var map = readStateMap();
            var s = map && map[tab];
            if (!s || typeof s !== 'object') return null;
            return s;
        }

        function setOptionHidden(opt, hidden) {
            if (!opt) return;
            opt.hidden = !!hidden;
            opt.disabled = !!hidden;
        }

        function normalizeTitleFilterByTab() {
            var v = (filterTitle.value || 'all').toLowerCase();

            if (state.tab === 'shared') {
                if (v !== 'all' && v !== 'viewed' && v !== 'unviewed') filterTitle.value = 'all';
            } else {
                if (v === 'viewed' || v === 'unviewed') filterTitle.value = 'all';
            }
            state.titleFilter = filterTitle.value;
        }

        function setSubWrapVisible(wrap, show) {
            if (!wrap) return;
            wrap.style.display = show ? 'flex' : 'none';
        }

        function applySubtabSelection(tabsArr, key, val) {
            if (!tabsArr || !tabsArr.length) return;
            tabsArr.forEach(function (b) {
                var sel = (String(b.dataset[key] || '') === String(val || ''));
                b.setAttribute('aria-selected', sel ? 'true' : 'false');
            });
        }

        function applySubtabsUi() {
            var isCreated = state.tab === 'created';
            var isApproval = state.tab === 'approval';
            var isCooperation = state.tab === 'cooperation';
            var isShared = state.tab === 'shared';

            setSubWrapVisible(createdSubWrap, isCreated);
            setSubWrapVisible(approvalSubWrap, isApproval);
            setSubWrapVisible(cooperationSubWrap, isCooperation);
            setSubWrapVisible(sharedSubWrap, isShared);

            if (isCreated) applySubtabSelection(createdSubTabs, 'createdSub', state.createdSub);
            if (isApproval) applySubtabSelection(approvalSubTabs, 'approvalSub', state.approvalSub);
            if (isCooperation) applySubtabSelection(cooperationSubTabs, 'cooperationSub', state.cooperationSub);
            if (isShared) applySubtabSelection(sharedSubTabs, 'sharedSub', state.sharedSub);
        }

        function applyTabUi() {
            var isShared = state.tab === 'shared';

            var sharedOnly = Array.from(filterTitle.querySelectorAll('option[data-shared-only="true"]'));
            var nonSharedOnly = Array.from(filterTitle.querySelectorAll('option[data-nonshared-only="true"]'));
            var createdOnly = Array.from(filterTitle.querySelectorAll('option[data-created-only="true"]'));

            sharedOnly.forEach(function (o) { setOptionHidden(o, !isShared); });
            nonSharedOnly.forEach(function (o) { setOptionHidden(o, isShared); });

            if (state.tab === 'created') createdOnly.forEach(function (o) { setOptionHidden(o, false); });
            else createdOnly.forEach(function (o) { setOptionHidden(o, true); });

            normalizeTitleFilterByTab();

            tabs.forEach(function (x) {
                x.setAttribute('aria-selected', x.dataset.tab === state.tab ? 'true' : 'false');
            });
            applySubtabsUi();
        }

        var state = {
            tab: 'created',
            page: 1,
            pageSize: 20,
            titleFilter: 'all',
            sort: 'created_desc',
            q: '',
            createdSub: 'ongoing',
            approvalSub: 'ongoing',
            cooperationSub: 'ongoing',
            sharedSub: 'ongoing'
        };

        function applyStateToControls() {
            try {
                filterTitle.value = state.titleFilter || 'all';
                filterSort.value = state.sort || 'created_desc';
                filterQuery.value = state.q || '';
            } catch { }
        }

        function buildStateKey() {
            var parts = [
                state.tab || '',
                state.page || 1,
                state.pageSize || 20,
                state.titleFilter || 'all',
                state.sort || 'created_desc',
                state.q || '',
                (state.tab === 'created' ? (state.createdSub || 'ongoing') : ''),
                (state.tab === 'approval' ? (state.approvalSub || 'ongoing') : ''),
                (state.tab === 'cooperation' ? (state.cooperationSub || 'ongoing') : ''),
                (state.tab === 'shared' ? (state.sharedSub || 'ongoing') : '')
            ];
            return parts.map(function (x) { return String(x); }).join('|');
        }

        function saveSnapshot(res) {
            try {
                var key = SNAPSHOT_PREFIX + buildStateKey();
                writeJson(key, { at: Date.now(), res: res || null });
            } catch { }
        }

        function tryRestoreSnapshot() {
            try {
                var key = SNAPSHOT_PREFIX + buildStateKey();
                var snap = readJson(key, null);
                if (!snap || !snap.res) return null;
                return snap.res;
            } catch { return null; }
        }

        function isApprovalOngoing() {
            return (state.tab === 'approval' && String(state.approvalSub || '') === 'ongoing');
        }

        function ensureBulkUiVisible() {
            var show = isApprovalOngoing();

            if (pagingTools) pagingTools.hidden = !show;

            if (show) {
                if (!chkAllTop) {
                    chkAllTop = document.createElement('input');
                    chkAllTop.type = 'checkbox';
                    chkAllTop.id = 'chkAllTop';
                    chkAllTop.addEventListener('change', function () {
                        setAllVisibleChecks(chkAllTop.checked);
                    });
                }

                if (noHeaderWrap) {
                    if (noHeaderWrap.dataset.bulkBound !== '1') {
                        var noText = noHeaderWrap.textContent || '';
                        noHeaderWrap.textContent = '';
                        var wrap = document.createElement('span');
                        wrap.className = 'doc-no-cell';
                        wrap.appendChild(chkAllTop);
                        var t = document.createElement('span');
                        t.textContent = noText;
                        wrap.appendChild(t);
                        noHeaderWrap.appendChild(wrap);
                        noHeaderWrap.dataset.bulkBound = '1';
                    }
                }

                if (chkAllBottom) {
                    chkAllBottom.onchange = function () { setAllVisibleChecks(chkAllBottom.checked); };
                }

                btnApproveBulk = findApproveButtonFallback();
                if (btnApproveBulk) {
                    try { btnApproveBulk.type = 'button'; } catch { }
                }
            } else {
                selectedDocIds.clear();
                if (chkAllBottom) chkAllBottom.checked = false;
                if (chkAllTop) chkAllTop.checked = false;
                updateBulkUiState();
            }
        }

        function setAllVisibleChecks(checked) {
            var rows = Array.from(body.querySelectorAll('tr[data-docid]'));
            rows.forEach(function (tr) {
                var docId = tr.dataset.docid;
                var cb = tr.querySelector('input[data-doccheck="1"]');
                if (!cb || !docId) return;
                cb.checked = !!checked;

                if (checked) selectedDocIds.add(String(docId));
                else selectedDocIds.delete(String(docId));
            });

            if (chkAllTop) chkAllTop.checked = !!checked;
            if (chkAllBottom) chkAllBottom.checked = !!checked;
            updateBulkUiState();
        }

        function updateBulkUiState() {
            var show = isApprovalOngoing();
            if (!show) {
                if (btnApproveBulk) btnApproveBulk.disabled = true;
                return;
            }

            var rows = Array.from(body.querySelectorAll('tr[data-docid]'));
            var total = rows.length;
            var checkedCount = 0;

            rows.forEach(function (tr) {
                var cb = tr.querySelector('input[data-doccheck="1"]');
                if (cb && cb.checked) checkedCount++;
            });

            var all = (total > 0 && checkedCount === total);
            if (chkAllTop) chkAllTop.checked = all;
            if (chkAllBottom) chkAllBottom.checked = all;

            if (btnApproveBulk) btnApproveBulk.disabled = (selectedDocIds.size === 0);
        }

        function renderList(res) {
            body.innerHTML = '';
            ensureBulkUiVisible();

            if (!res || !res.items || res.items.length === 0) {
                var tr = document.createElement('tr');
                var td = document.createElement('td');
                td.colSpan = getColSpan();
                td.textContent = LIST_EMPTY;
                tr.appendChild(td);
                body.appendChild(tr);
                paging.innerHTML = '';
                updateBulkUiState();
                return;
            }

            var total = Number(res.total || 0) || (res.page - 1) * res.pageSize + res.items.length;
            var startIndex = (res.page - 1) * res.pageSize;

            res.items.forEach(function (item, idx) {
                var displayNo = total - startIndex - idx;
                body.appendChild(buildRow(item, displayNo));
            });

            var totalPages = Math.max(1, Math.ceil((res.total || 0) / res.pageSize));
            paging.innerHTML = '';

            function makeBtn(label, page, current) {
                var b = document.createElement('button');
                b.textContent = label;
                if (current) b.setAttribute('aria-current', 'page');
                b.addEventListener('click', function () {
                    state.page = page;
                    selectedDocIds.clear();
                    setUrlFromState(true);
                    saveStateToSession();
                    loadList();
                });
                return b;
            }

            var block = 10;
            var blkIdx = Math.floor((state.page - 1) / block);
            var start = blkIdx * block + 1;
            var end = Math.min(totalPages, start + block - 1);

            if (start > 1) paging.appendChild(makeBtn(PAGINATION_PREV, start - 1, false));
            for (var p = start; p <= end; p++) paging.appendChild(makeBtn(String(p), p, p === state.page));
            if (end < totalPages) paging.appendChild(makeBtn(PAGINATION_NEXT, end + 1, false));

            updateBulkUiState();
        }

        function buildRow(item, displayNo) {
            var tr = document.createElement('tr');
            tr.dataset.docid = item.docId || item.DocId || '';

            tr.classList.remove('doc-row-hold');

            var rawStatusCode = String(item.statusCode || item.StatusCode || item.status || item.Status || '').trim();
            var baseStatus = getStatusBase(rawStatusCode);
            var statusUpper = rawStatusCode.toUpperCase();

            var isApprovalUnreadByStatus =
                (statusUpper.indexOf('PENDINGA') === 0) ||
                (statusUpper.indexOf('PENDINGHOLDA') === 0) ||
                (baseStatus === 'Pending') ||
                (baseStatus === 'OnHold');

            var rawIsRead = pick(item, ['isRead', 'IsRead', 'read', 'Read', 'is_viewed', 'isViewed', 'IsViewed']);
            var isRead = toBool(rawIsRead);
            var isUnread = !isRead;

            var shouldBold =
                (state.tab === 'approval' && isApprovalUnreadByStatus) ||
                ((state.tab === 'created' || state.tab === 'cooperation' || state.tab === 'shared') && isUnread);

            if (shouldBold) tr.classList.add('doc-row-unread');
            else tr.classList.remove('doc-row-unread');

            var tdNo = document.createElement('td');
            if (isApprovalOngoing()) {
                var cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.setAttribute('data-doccheck', '1');

                var docId = tr.dataset.docid || '';
                cb.checked = docId ? selectedDocIds.has(String(docId)) : false;

                cb.addEventListener('change', function () {
                    var id = tr.dataset.docid || '';
                    if (!id) return;
                    if (cb.checked) selectedDocIds.add(String(id));
                    else selectedDocIds.delete(String(id));
                    updateBulkUiState();
                });

                cb.style.marginRight = '10px';
                tdNo.appendChild(cb);
            }

            var noText = document.createElement('span');
            noText.textContent = String(displayNo);
            tdNo.appendChild(noText);

            var tdTitle = document.createElement('td');
            var wrapper = document.createElement('span');
            wrapper.className = 'doc-title';

            var link = document.createElement('a');
            var tabQs = (state && state.tab) ? ('?tab=' + encodeURIComponent(state.tab)) : '';
            link.href = detailUrl.replace('{docId}', (item.docId || item.DocId)) + tabQs;
            link.title = item.templateTitle || item.TemplateTitle || '';
            link.textContent = item.templateTitle || item.TemplateTitle || '';

            if (shouldBold) link.classList.add('doc-unread');
            else link.classList.remove('doc-unread');

            wrapper.appendChild(link);
            tdTitle.appendChild(wrapper);

            var tdAuthor = document.createElement('td');
            tdAuthor.textContent = item.authorName || item.AuthorName || '';

            var tdDate = document.createElement('td');
            tdDate.textContent = item.createdAt || item.CreatedAt || '';

            var tdStatus = document.createElement('td');
            var tdResult = document.createElement('td');

            var totalA = (typeof (item.totalApprovers ?? item.TotalApprovers) === 'number')
                ? (item.totalApprovers ?? item.TotalApprovers)
                : 0;

            var doneA = (typeof (item.completedApprovers ?? item.CompletedApprovers) === 'number')
                ? (item.completedApprovers ?? item.CompletedApprovers)
                : 0;

            var leftSummary = String((item.resultSummary || item.ResultSummary || '')).trim();
            var rightStatus = formatStatus(
                item.statusCode || item.StatusCode || item.status || item.Status,
                doneA,
                totalA
            );

            tdStatus.innerHTML =
                '<div class="doc-status-join" title="' + leftSummary.replace(/"/g, '&quot;') + ' ' + rightStatus.replace(/"/g, '&quot;') + '">' +
                '<span class="st-left"></span>' +
                '<span class="st-right"></span>' +
                '</div>';

            try {
                var join = tdStatus.querySelector('.doc-status-join');
                if (join) {
                    var l = join.querySelector('.st-left');
                    var r = join.querySelector('.st-right');
                    if (l) l.textContent = leftSummary;
                    if (r) r.textContent = rightStatus;
                } else {
                    tdStatus.textContent = (leftSummary ? (leftSummary + ' ') : '') + rightStatus;
                }
            } catch {
                tdStatus.textContent = (leftSummary ? (leftSummary + ' ') : '') + rightStatus;
            }

            tdResult.textContent = isRead ? READ_LABEL : UNREAD_LABEL;

            tr.append(tdNo, tdTitle, tdAuthor, tdDate, tdStatus, tdResult);
            return tr;
        }

        if (root && root.dataset && root.dataset.bulkApproveBound !== '1') {
            root.dataset.bulkApproveBound = '1';

            root.addEventListener('click', async function (ev) {
                if (!isApprovalOngoing()) return;

                var t = ev.target;
                if (!t) return;

                var btn = null;
                try { btn = t.closest('button'); } catch { btn = null; }
                if (!btn) return;

                var isBulkBtn =
                    (btn.id === 'btnApproveBulk') ||
                    (btn.getAttribute('data-bulk-approve') === '1') ||
                    (pagingTools && pagingTools.contains(btn) && ((btn.textContent || '').trim() === '승인'));

                if (!isBulkBtn) return;

                try { ev.preventDefault(); ev.stopPropagation(); } catch { }

                btnApproveBulk = btn;

                try {
                    var ids = Array.from(selectedDocIds.values())
                        .map(function (x) { return String(x || '').trim(); })
                        .filter(function (x) { return !!x; });

                    if (ids.length === 0) return;

                    btnApproveBulk.disabled = true;

                    var j = await doBulkApprove(ids);

                    selectedDocIds.clear();
                    if (chkAllTop) chkAllTop.checked = false;
                    if (chkAllBottom) chkAllBottom.checked = false;

                    await refreshBadges();
                    await loadList();

                    var approved = 0;
                    try {
                        approved = asInt(j && j.totals ? (j.totals.approved ?? j.totals.Approved) : 0);
                    } catch { approved = 0; }

                    if (!(approved >= 0) || (j && j.ok === true && !(j && j.totals && ('approved' in j.totals || 'Approved' in j.totals)))) {
                        console.error('[DocBoard] BulkApprove ok but missing totals.approved:', j);
                        if (BULK_FAIL) showTopBarToast(BULK_FAIL, 'danger');
                        return;
                    }

                    if (BULK_OK_FMT) showTopBarToast(fmt0(BULK_OK_FMT, approved), 'success');

                    console.debug('[DocBoard] BulkApprove result:', j);
                } catch (e) {
                    console.error('[DocBoard] BulkApprove failed:', e && e.payload ? e.payload : e);
                    if (BULK_FAIL) showTopBarToast(BULK_FAIL, 'danger');
                } finally {
                    updateBulkUiState();
                }
            }, true);
        }

        async function loadList() {
            try {
                body.innerHTML = '<tr><td colspan="' + getColSpan() + '">' + LIST_LOADING + '</td></tr>';

                var params = new URLSearchParams(window.location.search || '');
                params.set('tab', state.tab || 'created');
                params.set('page', String(state.page || 1));
                params.set('pageSize', String(state.pageSize || 20));
                params.set('titleFilter', String(state.titleFilter || 'all'));
                params.set('sort', String(state.sort || 'created_desc'));
                params.set('q', String(state.q || ''));

                if (state.tab === 'created') params.set('createdSub', String(state.createdSub || 'ongoing'));
                else params.delete('createdSub');

                if (state.tab === 'approval') params.set('approvalSub', String(state.approvalSub || 'ongoing'));
                else params.delete('approvalSub');

                if (state.tab === 'cooperation') params.set('cooperationSub', String(state.cooperationSub || 'ongoing'));
                else params.delete('cooperationSub');

                if (state.tab === 'shared') params.set('sharedSub', String(state.sharedSub || 'ongoing'));
                else params.delete('sharedSub');

                params.set('_ts', String(Date.now()));

                var url = apiList + (apiList.indexOf('?') >= 0 ? '&' : '?') + params.toString();
                console.debug('[DocBoard] BoardData fetch:', url);

                var r = await fetch(url, {
                    method: 'GET',
                    cache: 'no-store',
                    headers: { 'Accept': 'application/json' }
                });

                var res = await r.json().catch(function () { return null; });
                if (!r.ok || !res) {
                    console.error('[DocBoard] BoardData failed:', r.status, res);
                    body.innerHTML = '<tr><td colspan="' + getColSpan() + '">' + LIST_EMPTY + '</td></tr>';
                    paging.innerHTML = '';
                    return;
                }

                saveSnapshot(res);
                renderList(res);
            } catch (e) {
                console.error('[DocBoard] loadList failed:', e);
                try {
                    body.innerHTML = '<tr><td colspan="' + getColSpan() + '">' + LIST_EMPTY + '</td></tr>';
                    paging.innerHTML = '';
                } catch { }
            }
        }

        async function refreshBadges() {
            if (!apiBadges) return;
            try {
                var url = apiBadges + (apiBadges.indexOf('?') >= 0 ? '&' : '?') + '_ts=' + Date.now();
                var res = await fetch(url, { method: 'GET', cache: 'no-store', headers: { 'Accept': 'application/json' } });
                var j = await res.json().catch(function () { return null; });
                if (!res.ok || !j) return;

                var createdUnread = (j.createdUnread ?? j.CreatedUnread ?? j.created ?? j.Created ?? 0);
                var approvalPending = (j.approvalPending ?? j.ApprovalPending ?? j.approval ?? j.Approval ?? 0);
                var cooperationPending = (j.cooperationPending ?? j.CooperationPending ?? j.cooperation ?? j.Cooperation ?? 0);
                var sharedUnread = (j.sharedUnread ?? j.SharedUnread ?? j.shared ?? j.Shared ?? 0);

                setBadgeDom('badge-created', createdUnread);
                setBadgeDom('badge-approval', approvalPending);
                setBadgeDom('badge-cooperation', cooperationPending);
                setBadgeDom('badge-shared', sharedUnread);
            } catch { }
        }

        function restoreInitialState() {
            var fromUrl = readStateFromUrl();

            var tab = fromUrl.tab || null;
            if (!tab) {
                try {
                    var storedTab = window.sessionStorage.getItem('docBoardTab');
                    if (storedTab && allowedTabs.indexOf(storedTab) >= 0) tab = storedTab;
                } catch { }
            }
            if (!tab) tab = 'created';
            state.tab = tab;

            var fromSession = readStateForTabFromSession(state.tab) || {};

            state.page = fromUrl.page || fromSession.page || 1;
            state.pageSize = fromUrl.pageSize || fromSession.pageSize || 20;
            state.titleFilter = fromUrl.titleFilter || fromSession.titleFilter || 'all';
            state.sort = fromUrl.sort || fromSession.sort || 'created_desc';
            state.q = (fromUrl.q != null && fromUrl.q !== '') ? fromUrl.q : (fromSession.q || '');

            state.createdSub = (fromUrl.createdSub || fromSession.createdSub || 'ongoing');
            state.approvalSub = (fromUrl.approvalSub || fromSession.approvalSub || 'ongoing');
            state.cooperationSub = (fromUrl.cooperationSub || fromSession.cooperationSub || 'ongoing');
            state.sharedSub = (fromUrl.sharedSub || fromSession.sharedSub || 'ongoing');

            applyStateToControls();
            applyTabUi();

            setUrlFromState(true);
            saveStateToSession();

            try { window.sessionStorage.setItem('docBoardTab', state.tab); } catch { }

            ensureBulkUiVisible();
        }

        function bindSubtabs(tabKey, tabsArr, valueKey) {
            (tabsArr || []).forEach(function (b) {
                b.addEventListener('click', function () {
                    if (state.tab !== tabKey) return;

                    var v = b.dataset[valueKey];
                    if (!v) return;

                    state[valueKey] = v;
                    state.page = 1;

                    selectedDocIds.clear();

                    applySubtabsUi();
                    setUrlFromState(true);
                    saveStateToSession();

                    ensureBulkUiVisible();
                    loadList();
                });
            });
        }

        bindSubtabs('created', createdSubTabs, 'createdSub');
        bindSubtabs('approval', approvalSubTabs, 'approvalSub');
        bindSubtabs('cooperation', cooperationSubTabs, 'cooperationSub');
        bindSubtabs('shared', sharedSubTabs, 'sharedSub');

        tabs.forEach(function (t) {
            t.addEventListener('click', function () {
                var nextTab = t.dataset.tab;
                if (!nextTab || allowedTabs.indexOf(nextTab) < 0) return;

                state.tab = nextTab;

                var s = readStateForTabFromSession(state.tab);
                if (s) {
                    state.page = s.page || 1;
                    state.pageSize = s.pageSize || 20;
                    state.titleFilter = s.titleFilter || 'all';
                    state.sort = s.sort || 'created_desc';
                    state.q = s.q || '';
                    state.createdSub = s.createdSub || state.createdSub || 'ongoing';
                    state.approvalSub = s.approvalSub || state.approvalSub || 'ongoing';
                    state.cooperationSub = s.cooperationSub || state.cooperationSub || 'ongoing';
                    state.sharedSub = s.sharedSub || state.sharedSub || 'ongoing';
                } else {
                    state.page = 1;
                    state.titleFilter = filterTitle.value || 'all';
                    state.sort = filterSort.value || 'created_desc';
                    state.q = filterQuery.value || '';
                }

                selectedDocIds.clear();

                applyStateToControls();
                applyTabUi();

                try { window.sessionStorage.setItem('docBoardTab', state.tab); } catch { }
                setUrlFromState(true);
                saveStateToSession();

                ensureBulkUiVisible();
                refreshBadges();
                loadList();
            });
        });

        if (filterQuery) {
            filterQuery.addEventListener('keydown', function (e) {
                if (e.key === 'Enter' && btnApply) btnApply.click();
            });
        }

        if (btnApply) {
            btnApply.addEventListener('click', function () {
                state.titleFilter = filterTitle.value;
                state.sort = filterSort.value;
                state.q = filterQuery.value || '';
                state.pageSize = clamp(state.pageSize, 1, 100);
                state.page = 1;

                selectedDocIds.clear();

                applyStateToControls();
                applyTabUi();

                setUrlFromState(true);
                saveStateToSession();

                ensureBulkUiVisible();
                loadList();
            });
        }

        window.addEventListener('popstate', function () {
            var u = readStateFromUrl();
            if (u.tab && allowedTabs.indexOf(u.tab) >= 0) state.tab = u.tab;
            state.page = u.page || 1;
            state.pageSize = u.pageSize || 20;
            state.titleFilter = u.titleFilter || 'all';
            state.sort = u.sort || 'created_desc';
            state.q = (u.q != null) ? u.q : '';
            if (u.createdSub) state.createdSub = u.createdSub;
            if (u.approvalSub) state.approvalSub = u.approvalSub;
            if (u.cooperationSub) state.cooperationSub = u.cooperationSub;
            if (u.sharedSub) state.sharedSub = u.sharedSub;

            selectedDocIds.clear();
            applyStateToControls();
            applyTabUi();
            saveStateToSession();

            var snap = tryRestoreSnapshot();
            if (snap) renderList(snap);

            loadList();
            refreshBadges();
        });

        restoreInitialState();

        var snap = tryRestoreSnapshot();
        if (snap) renderList(snap);

        loadList();
        refreshBadges();
    }

    window.EBDocBoard = window.EBDocBoard || {};
    window.EBDocBoard.init = function (rootEl) {
        if (!rootEl) return;
        try {
            if (rootEl.dataset && rootEl.dataset._boardInit === '1') return;
            if (rootEl.dataset) rootEl.dataset._boardInit = '1';
        } catch { }
        createBoard(rootEl);
    };

    window.DocBoard = window.DocBoard || {};
    window.DocBoard.init = function (rootEl) {
        return window.EBDocBoard.init(rootEl);
    };

    document.addEventListener('DOMContentLoaded', function () {
        var root = document.getElementById('docBoard') || document.querySelector('[data-api-list]');
        if (!root) {
            console.warn('[DocBoard] root not found for auto init');
            return;
        }
        window.EBDocBoard.init(root);
    });
})();