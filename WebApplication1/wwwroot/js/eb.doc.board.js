// 2026.03.31 Changed: 상태 컬럼을 도트+이름 방식으로 교체
//   - 진행중: 도트(완료=초록/반려=빨강/회수=검정/대기=연파랑) + 현재 대기자 이름·직급
//   - 승인완료: 도트 전체 초록 + "승인됨"
//   - 반려됨: 반려 도트 + "반려됨"
//   - 회수됨: 회수 도트 + "회수됨"
//   - 협조선 있으면 구분선(|) 뒤에 동일 패턴으로 병기
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
            var el = document.getElementById('docBoardI18n');
            if (!el) return {};
            var raw = (el.textContent || '').trim();
            if (!raw) return {};
            var obj = JSON.parse(raw);
            return (obj && typeof obj === 'object') ? obj : {};
        } catch { return {}; }
    }

    function toBool(v) {
        if (v === true) return true;
        if (v === false) return false;
        if (v == null) return false;
        if (typeof v === 'number') return v !== 0;
        if (typeof v === 'string') {
            var s = v.trim().toLowerCase();
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
            btn.addEventListener('click', function () { try { box.remove(); } catch { } });
            box.appendChild(msg);
            box.appendChild(btn);
            host.appendChild(box);
            setTimeout(function () { try { box.remove(); } catch { } }, 3500);
        } catch { }
    }

    // ─────────────────────────────────────────────────────────
    // 도트 상태 렌더링 헬퍼
    // ─────────────────────────────────────────────────────────
    var DOT_STYLES = {
        done: 'background:#4a8fdd;',   // 완료·승인 — 파랑
        approve: 'background:#4a8fdd;',   // 최종승인  — 파랑
        reject: 'background:#e05252;',   // 반려      — 빨강
        recall: 'background:#2c2c2c;',   // 회수      — 검정
        collab: 'background:#c0c0c0;',   // 협조 현재대기 — 회색
        wait: 'background:#c0c0c0;',   // 결재 대기중   — 회색
        todo: 'background:#c0c0c0;'    // 미도달        — 회색
    };

    var DOT_BASE = 'display:inline-block;width:10px;height:10px;border-radius:50%;flex-shrink:0;';

    function makeDot(type) {
        var span = document.createElement('span');
        span.setAttribute('style', DOT_BASE + (DOT_STYLES[type] || DOT_STYLES.todo));
        return span;
    }

    function makeDots(total, done, currentDotType) {
        // total: 전체 단계 수
        // done: 완료(초록) 단계 수  — currentDotType 도트 포함 안 함
        // currentDotType: 현재 단계 도트 색상 키
        var wrap = document.createElement('span');
        wrap.setAttribute('style', 'display:inline-flex;gap:3px;align-items:center;flex-shrink:0;');

        for (var i = 0; i < total; i++) {
            if (i < done) {
                wrap.appendChild(makeDot('done'));
            } else if (i === done) {
                wrap.appendChild(makeDot(currentDotType));
            } else {
                wrap.appendChild(makeDot('todo'));
            }
        }
        return wrap;
    }

    function makeDivider() {
        var span = document.createElement('span');
        span.setAttribute('style', 'width:.5px;height:10px;background:#d1d1d1;margin:0 5px;flex-shrink:0;');
        return span;
    }

    function makeNameSpan(name, position) {
        var span = document.createElement('span');
        span.setAttribute('style', 'font-size:15px;color:inherit;white-space:nowrap;');
        span.textContent = name || '';
        if (position) {
            var pos = document.createElement('span');
            pos.setAttribute('style', 'font-size:15px;color:inherit;margin-left:3px;');
            pos.textContent = position;
            span.appendChild(pos);
        }
        return span;
    }

    function makeLabelSpan(text, color) {
        var span = document.createElement('span');
        span.setAttribute('style', 'font-size:15px;font-weight:500;color:' + color + ';white-space:nowrap;');
        span.textContent = text || '';
        return span;
    }

    // i18n 라벨 — 기존 STATUS_LABELS 외 도트 전용 라벨
    // View의 docBoardI18n 에 아래 키를 추가하면 다국어 대응됨
    // DOC_Status_ApprovedDone / DOC_Status_RejectedDone / DOC_Status_RecalledDone
    // 없으면 아래 기본값 사용
    function getDotLabel(key, i18n) {
        var labels = (i18n && i18n.DOT_LABELS) ? i18n.DOT_LABELS : {};
        var defaults = { ApprovedDone: '승인됨', RejectedDone: '반려됨', RecalledDone: '회수됨' };
        return labels[key] || defaults[key] || key;
    }

    /**
     * buildStatusCell(item, i18n)
     * 결재선 도트 + 협조선 도트(있을 경우)를 DOM으로 만들어 반환
     */
    function buildStatusCell(item, i18n) {
        var wrap = document.createElement('div');
        wrap.setAttribute('style',
            'display:inline-flex;align-items:center;gap:5px;white-space:nowrap;');

        var rawStatus = String(
            item.statusCode || item.StatusCode ||
            item.status || item.Status || ''
        ).trim();

        var base = getStatusBase(rawStatus);

        var totalA = asInt(item.totalApprovers ?? item.TotalApprovers ?? 0);
        var doneA = asInt(item.completedApprovers ?? item.CompletedApprovers ?? 0);
        var actorName = String(item.resultSummary || item.ResultSummary || '').trim();
        // resultSummary 는 "이름 직급" 형태로 이미 조합돼서 내려옴
        // 도트 옆 이름 표시에 그대로 사용

        // ── 결재선 도트 결정 ──────────────────────────────
        var apDotType, apLabel, apLabelColor;

        if (base === 'Approved') {
            // 전체 승인 완료 — 모든 도트 파랑, "승인됨" 텍스트
            apDotType = 'approve';
            apLabel = getDotLabel('ApprovedDone', i18n);
            apLabelColor = '#4a8fdd';
            doneA = totalA;
        } else if (base === 'Rejected') {
            apDotType = 'reject';
            apLabel = getDotLabel('RejectedDone', i18n);
            apLabelColor = '#e05252';
        } else if (base === 'Recalled') {
            apDotType = 'recall';
            apLabel = getDotLabel('RecalledDone', i18n);
            apLabelColor = '#2c2c2c';
        } else {
            // Pending / OnHold — 진행중, 이름만 표시
            apDotType = 'wait';
            apLabel = null;
            apLabelColor = null;
        }

        // 결재선 렌더
        if (totalA > 0) {
            wrap.appendChild(makeDots(totalA, doneA, apDotType));
        }

        if (apLabel) {
            wrap.appendChild(makeLabelSpan(apLabel, apLabelColor));
        } else if (actorName) {
            var nameSpan = document.createElement('span');
            nameSpan.setAttribute('style', 'font-size:15px;color:inherit;white-space:nowrap;');
            nameSpan.textContent = actorName;
            wrap.appendChild(nameSpan);
        }

        // ── 협조선 도트 (coopTotalSteps > 0 인 경우만) ──
        var coopTotal = asInt(item.coopTotalSteps ?? item.CoopTotalSteps ?? 0);
        var coopDoneKeys = new Set(
            String(item.coopDoneKeys ?? item.CoopDoneKeys ?? '').split(',').filter(Boolean).map(Number)
        );
        var coopName = String(item.coopPendingName ?? item.CoopPendingName ?? '').trim();
        var coopPos = String(item.coopPendingPosition ?? item.CoopPendingPosition ?? '').trim();

        if (coopTotal > 0) {
            wrap.appendChild(makeDivider());

            if (coopDoneKeys.size >= coopTotal) {
                // 협조 전체 완료
                var coopWrap = document.createElement('span');
                coopWrap.setAttribute('style', 'display:inline-flex;gap:3px;align-items:center;flex-shrink:0;');
                for (var i = 1; i <= coopTotal; i++) {
                    coopWrap.appendChild(makeDot('approve'));
                }
                wrap.appendChild(coopWrap);
                wrap.appendChild(makeLabelSpan(getDotLabel('ApprovedDone', i18n), '#4a8fdd'));
            } else {
                // 협조 진행중 — RoleKey 번호 위치에 파랑/회색
                var coopWrap = document.createElement('span');
                coopWrap.setAttribute('style', 'display:inline-flex;gap:3px;align-items:center;flex-shrink:0;');
                for (var i = 1; i <= coopTotal; i++) {
                    coopWrap.appendChild(makeDot(coopDoneKeys.has(i) ? 'approve' : 'collab'));
                }
                wrap.appendChild(coopWrap);
                if (coopName) {
                    wrap.appendChild(makeNameSpan(coopName, coopPos));
                }
            }
        }
        return wrap;
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

    // ─────────────────────────────────────────────────────────

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
            document.getElementById('docListBody') ||
            $('tbody#docListBody', root) ||
            $('tbody[data-doc-list-body="1"]', root) ||
            $('table tbody', root);

        var paging =
            document.getElementById('docPaging') ||
            $('#docPaging', root) ||
            $('[data-doc-paging="1"]', root);

        if (!body) {
            console.error('[DocBoard] list tbody not found.');
            return;
        }
        if (!paging) {
            console.warn('[DocBoard] paging element not found.');
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
                    if ((btns[i].textContent || '').trim() === '승인') return btns[i];
                }
            }
            return document.querySelector('button[data-bulk-approve="1"]') || null;
        }

        var chkAllTop = null;
        var selectedDocIds = new Set();

        var clamp = function (v, mn, mx) { return Math.max(mn, Math.min(mx, v)); };
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

        function getColSpan() { return isApprovalOngoing() ? 7 : 6; }

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
            var res = await fetch('/Doc/BulkApprove', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'RequestVerificationToken': token
                },
                credentials: 'same-origin',
                body: JSON.stringify({ DocIds: DocIds || [] })
            });
            var j = await res.json().catch(function () { return null; });
            if (!res.ok || !j || j.ok !== true) {
                var err = new Error('bulk approve failed');
                err.payload = j; err.status = res.status;
                throw err;
            }
            return j;
        }

        function readJson(key, fallback) {
            try {
                var s = window.sessionStorage.getItem(key);
                if (!s) return fallback;
                var o = JSON.parse(s);
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
                tab: state.tab, page: state.page, pageSize: state.pageSize,
                titleFilter: state.titleFilter, sort: state.sort, q: state.q,
                createdSub: state.createdSub, approvalSub: state.approvalSub,
                cooperationSub: state.cooperationSub, sharedSub: state.sharedSub
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

                if (state.tab === 'created') { qs.set('createdSub', String(state.createdSub || 'ongoing')); } else qs.delete('createdSub');
                if (state.tab === 'approval') { qs.set('approvalSub', String(state.approvalSub || 'ongoing')); } else qs.delete('approvalSub');
                if (state.tab === 'cooperation') { qs.set('cooperationSub', String(state.cooperationSub || 'ongoing')); } else qs.delete('cooperationSub');
                if (state.tab === 'shared') { qs.set('sharedSub', String(state.sharedSub || 'ongoing')); } else qs.delete('sharedSub');

                var url = window.location.pathname + '?' + qs.toString();
                if (replaceOnly) window.history.replaceState(null, '', url);
                else window.history.pushState(null, '', url);
            } catch { }
        }

        function readStateFromUrl() {
            var qs = new URLSearchParams(window.location.search || '');
            var out = {};
            var tab = qs.get('tab');
            if (tab && allowedTabs.indexOf(tab) >= 0) out.tab = tab;
            var p = parseInt(qs.get('page') || '', 10);
            var ps = parseInt(qs.get('pageSize') || '', 10);
            if (!isNaN(p) && p > 0) out.page = p;
            if (!isNaN(ps) && ps > 0 && ps <= 100) out.pageSize = ps;
            var tf = qs.get('titleFilter'); if (tf) out.titleFilter = tf;
            var so = qs.get('sort'); if (so) out.sort = so;
            var q = qs.get('q'); if (q != null) out.q = q;
            var cs = qs.get('createdSub'); if (cs) out.createdSub = cs;
            var as = qs.get('approvalSub'); if (as) out.approvalSub = as;
            var cos = qs.get('cooperationSub'); if (cos) out.cooperationSub = cos;
            var ss = qs.get('sharedSub'); if (ss) out.sharedSub = ss;
            return out;
        }

        function readStateForTabFromSession(tab) {
            var map = readStateMap();
            var s = map && map[tab];
            return (s && typeof s === 'object') ? s : null;
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
                b.setAttribute('aria-selected', String(b.dataset[key] || '') === String(val || '') ? 'true' : 'false');
            });
        }

        function applySubtabsUi() {
            setSubWrapVisible(createdSubWrap, state.tab === 'created');
            setSubWrapVisible(approvalSubWrap, state.tab === 'approval');
            setSubWrapVisible(cooperationSubWrap, state.tab === 'cooperation');
            setSubWrapVisible(sharedSubWrap, state.tab === 'shared');

            if (state.tab === 'created') applySubtabSelection(createdSubTabs, 'createdSub', state.createdSub);
            if (state.tab === 'approval') applySubtabSelection(approvalSubTabs, 'approvalSub', state.approvalSub);
            if (state.tab === 'cooperation') applySubtabSelection(cooperationSubTabs, 'cooperationSub', state.cooperationSub);
            if (state.tab === 'shared') applySubtabSelection(sharedSubTabs, 'sharedSub', state.sharedSub);
        }

        function applyTabUi() {
            var isShared = state.tab === 'shared';
            Array.from(filterTitle.querySelectorAll('option[data-shared-only="true"]'))
                .forEach(function (o) { setOptionHidden(o, !isShared); });
            Array.from(filterTitle.querySelectorAll('option[data-nonshared-only="true"]'))
                .forEach(function (o) { setOptionHidden(o, isShared); });
            var createdOnly = Array.from(filterTitle.querySelectorAll('option[data-created-only="true"]'));
            createdOnly.forEach(function (o) { setOptionHidden(o, state.tab !== 'created'); });

            normalizeTitleFilterByTab();
            tabs.forEach(function (x) {
                x.setAttribute('aria-selected', x.dataset.tab === state.tab ? 'true' : 'false');
            });
            applySubtabsUi();
        }

        var state = {
            tab: 'created', page: 1, pageSize: 20,
            titleFilter: 'all', sort: 'created_desc', q: '',
            createdSub: 'ongoing', approvalSub: 'ongoing',
            cooperationSub: 'ongoing', sharedSub: 'ongoing'
        };

        function applyStateToControls() {
            try {
                filterTitle.value = state.titleFilter || 'all';
                filterSort.value = state.sort || 'created_desc';
                filterQuery.value = state.q || '';
            } catch { }
        }

        function buildStateKey() {
            return [
                state.tab, state.page, state.pageSize, state.titleFilter, state.sort, state.q,
                state.tab === 'created' ? (state.createdSub || 'ongoing') : '',
                state.tab === 'approval' ? (state.approvalSub || 'ongoing') : '',
                state.tab === 'cooperation' ? (state.cooperationSub || 'ongoing') : '',
                state.tab === 'shared' ? (state.sharedSub || 'ongoing') : ''
            ].map(String).join('|');
        }

        function saveSnapshot(res) {
            try { writeJson(SNAPSHOT_PREFIX + buildStateKey(), { at: Date.now(), res: res || null }); } catch { }
        }

        function tryRestoreSnapshot() {
            try {
                var snap = readJson(SNAPSHOT_PREFIX + buildStateKey(), null);
                return (snap && snap.res) ? snap.res : null;
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
                    chkAllTop.addEventListener('change', function () { setAllVisibleChecks(chkAllTop.checked); });
                }
                if (noHeaderWrap && noHeaderWrap.dataset.bulkBound !== '1') {
                    var noText = noHeaderWrap.textContent || '';
                    noHeaderWrap.textContent = '';
                    var wrapEl = document.createElement('span');
                    wrapEl.className = 'doc-no-cell';
                    wrapEl.appendChild(chkAllTop);
                    var t = document.createElement('span');
                    t.textContent = noText;
                    wrapEl.appendChild(t);
                    noHeaderWrap.appendChild(wrapEl);
                    noHeaderWrap.dataset.bulkBound = '1';
                }
                if (chkAllBottom) chkAllBottom.onchange = function () { setAllVisibleChecks(chkAllBottom.checked); };
                btnApproveBulk = findApproveButtonFallback();
                if (btnApproveBulk) { try { btnApproveBulk.type = 'button'; } catch { } }
            } else {
                selectedDocIds.clear();
                if (chkAllBottom) chkAllBottom.checked = false;
                if (chkAllTop) chkAllTop.checked = false;
                updateBulkUiState();
            }
        }

        function setAllVisibleChecks(checked) {
            Array.from(body.querySelectorAll('tr[data-docid]')).forEach(function (tr) {
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
            if (!isApprovalOngoing()) {
                if (btnApproveBulk) btnApproveBulk.disabled = true;
                return;
            }
            var rows = Array.from(body.querySelectorAll('tr[data-docid]'));
            var total = rows.length;
            var checked = 0;
            rows.forEach(function (tr) {
                var cb = tr.querySelector('input[data-doccheck="1"]');
                if (cb && cb.checked) checked++;
            });
            var all = (total > 0 && checked === total);
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
                body.appendChild(buildRow(item, total - startIndex - idx));
            });

            var totalPages = Math.max(1, Math.ceil((res.total || 0) / res.pageSize));
            paging.innerHTML = '';

            function makeBtn(label, pg, current) {
                var b = document.createElement('button');
                b.textContent = label;
                if (current) b.setAttribute('aria-current', 'page');
                b.addEventListener('click', function () {
                    state.page = pg;
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

            var rawStatusCode = String(
                item.statusCode || item.StatusCode ||
                item.status || item.Status || ''
            ).trim();
            var baseStatus = getStatusBase(rawStatusCode);
            var statusUpper = rawStatusCode.toUpperCase();

            var isApprovalUnreadByStatus =
                (statusUpper.indexOf('PENDINGA') === 0) ||
                (statusUpper.indexOf('PENDINGHOLDA') === 0) ||
                (baseStatus === 'Pending') ||
                (baseStatus === 'OnHold');

            var rawIsRead = pick(item, ['isRead', 'IsRead', 'read', 'Read', 'is_viewed', 'isViewed', 'IsViewed']);
            var isRead = toBool(rawIsRead);
            var shouldBold =
                (state.tab === 'approval' && isApprovalUnreadByStatus) ||
                ((state.tab === 'created' || state.tab === 'cooperation' || state.tab === 'shared') && !isRead);

            if (shouldBold) tr.classList.add('doc-row-unread');
            else tr.classList.remove('doc-row-unread');

            // ── No 칸 ──────────────────────────────────────
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

            // ── 제목 칸 ────────────────────────────────────
            var tdTitle = document.createElement('td');
            var wrapper = document.createElement('span');
            wrapper.className = 'doc-title';
            var link = document.createElement('a');
            var tabQs = (state && state.tab) ? ('?tab=' + encodeURIComponent(state.tab)) : '';
            link.href = detailUrl.replace('{docId}', (item.docId || item.DocId)) + tabQs;
            link.title = item.templateTitle || item.TemplateTitle || '';
            link.textContent = item.templateTitle || item.TemplateTitle || '';
            if (shouldBold) link.classList.add('doc-unread');
            wrapper.appendChild(link);
            tdTitle.appendChild(wrapper);

            // ── 작성자 / 일시 ──────────────────────────────
            var tdAuthor = document.createElement('td');
            tdAuthor.textContent = item.authorName || item.AuthorName || '';

            var tdDate = document.createElement('td');
            tdDate.textContent = item.createdAt || item.CreatedAt || '';

            // ── 상태 칸 (도트 렌더링) ──────────────────────
            var tdStatus = document.createElement('td');
            tdStatus.style.textAlign = 'center';
            tdStatus.appendChild(buildStatusCell(item, i18n));

            // ── 열람 칸 ────────────────────────────────────
            var tdResult = document.createElement('td');
            tdResult.textContent = isRead ? READ_LABEL : UNREAD_LABEL;

            tr.append(tdNo, tdTitle, tdAuthor, tdDate, tdStatus, tdResult);
            return tr;
        }

        // ── 일괄 승인 이벤트 ──────────────────────────────
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
                    try { approved = asInt(j && j.totals ? (j.totals.approved ?? j.totals.Approved) : 0); } catch { approved = 0; }
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
                if (state.tab === 'created') { params.set('createdSub', String(state.createdSub || 'ongoing')); } else params.delete('createdSub');
                if (state.tab === 'approval') { params.set('approvalSub', String(state.approvalSub || 'ongoing')); } else params.delete('approvalSub');
                if (state.tab === 'cooperation') { params.set('cooperationSub', String(state.cooperationSub || 'ongoing')); } else params.delete('cooperationSub');
                if (state.tab === 'shared') { params.set('sharedSub', String(state.sharedSub || 'ongoing')); } else params.delete('sharedSub');
                params.set('_ts', String(Date.now()));
                var url = apiList + (apiList.indexOf('?') >= 0 ? '&' : '?') + params.toString();
                console.debug('[DocBoard] BoardData fetch:', url);
                var r = await fetch(url, { method: 'GET', cache: 'no-store', headers: { 'Accept': 'application/json' } });
                var res = await r.json().catch(function () { return null; });
                if (!r.ok || !res) {
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
                setBadgeDom('badge-created', j.createdUnread ?? j.CreatedUnread ?? j.created ?? j.Created ?? 0);
                setBadgeDom('badge-approval', j.approvalPending ?? j.ApprovalPending ?? j.approval ?? j.Approval ?? 0);
                setBadgeDom('badge-cooperation', j.cooperationPending ?? j.CooperationPending ?? j.cooperation ?? j.Cooperation ?? 0);
                setBadgeDom('badge-shared', j.sharedUnread ?? j.SharedUnread ?? j.shared ?? j.Shared ?? 0);
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
            state.createdSub = fromUrl.createdSub || fromSession.createdSub || 'ongoing';
            state.approvalSub = fromUrl.approvalSub || fromSession.approvalSub || 'ongoing';
            state.cooperationSub = fromUrl.cooperationSub || fromSession.cooperationSub || 'ongoing';
            state.sharedSub = fromUrl.sharedSub || fromSession.sharedSub || 'ongoing';

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
                    state.page = s.page || 1; state.pageSize = s.pageSize || 20;
                    state.titleFilter = s.titleFilter || 'all'; state.sort = s.sort || 'created_desc';
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
    window.DocBoard.init = function (rootEl) { return window.EBDocBoard.init(rootEl); };

    document.addEventListener('DOMContentLoaded', function () {
        var root = document.getElementById('docBoard') || document.querySelector('[data-api-list]');
        if (!root) { console.warn('[DocBoard] root not found for auto init'); return; }
        window.EBDocBoard.init(root);
    });
})();
