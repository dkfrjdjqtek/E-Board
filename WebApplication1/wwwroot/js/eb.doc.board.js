// 2026.06.09 Changed: Board 작성일 표시/필터를 DocManage/DocTemplateManage와 동일하게 DocControllerHelper 공용 날짜 헬퍼 기준으로 통일
(function () {
    'use strict';

    var __docBoardCurrentUserId = '';

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
        } catch {
            return {};
        }
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

    function norm(v) {
        return String(v ?? '').trim().toUpperCase();
    }

    function normalizeId(v) {
        return String(v ?? '').trim().toLowerCase();
    }

    function fmt0(template, n) {
        try { return String(template || '').replace('{0}', String(n)); } catch { return String(n); }
    }

    function addUnique(parts, value) {
        var s = String(value || '').trim();
        if (!s) return;
        if (parts.indexOf(s) < 0) parts.push(s);
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

    var DOT_STYLES = {
        done: 'background:#4a8fdd;',
        approve: 'background:#4a8fdd;',
        reject: 'background:#e05252;',
        veto: 'background:#e05252;',
        recall: 'background:#2c2c2c;',
        hold: 'background:#f2c94c;',
        collab: 'background:#c0c0c0;',
        wait: 'background:#c0c0c0;',
        todo: 'background:#c0c0c0;'
    };

    var DOT_BASE = 'display:inline-block;width:10px;height:10px;border-radius:50%;flex-shrink:0;';

    function makeDot(type) {
        var span = document.createElement('span');
        span.setAttribute('style', DOT_BASE + (DOT_STYLES[type] || DOT_STYLES.todo));
        return span;
    }

    function makeDivider() {
        var span = document.createElement('span');
        span.setAttribute('style', 'width:1px;height:12px;background:#d1d1d1;margin:0 2px;flex-shrink:0;');
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

    function makeLabelSpan(text) {
        var span = document.createElement('span');
        span.setAttribute('style', 'font-size:15px;color:inherit;white-space:nowrap;');
        span.textContent = text || '';
        return span;
    }

    function getDotLabel(key, i18n) {
        var labels = (i18n && i18n.DOT_LABELS) ? i18n.DOT_LABELS : {};
        var defaults = {
            ApprovedDone: '승인됨',
            RejectedDone: '반려됨',
            VetoedDone: '부결됨',
            RecalledDone: '회수됨',
            OnHoldDone: '보류',
            Pending: '대기'
        };

        return labels[key] || defaults[key] || key;
    }

    function getStatusBase(raw) {
        if (!raw) return 'Pending';

        var s = String(raw).trim().toUpperCase();

        if (s.indexOf('PENDINGHOLD') === 0) return 'OnHold';
        if (s === 'HOLD' || s.indexOf('ONHOLD') === 0 || s.indexOf('ON HOLD') === 0) return 'OnHold';
        if (s.indexOf('PENDING') === 0) return 'Pending';
        if (s === 'APPROVE' || s.indexOf('APPROVED') === 0) return 'Approved';
        if (s === 'REJECT' || s.indexOf('REJECTED') === 0) return 'Rejected';
        if (s === 'VETO' || s.indexOf('VETOED') === 0) return 'Vetoed';
        if (s === 'RECALL' || s.indexOf('RECALLED') === 0) return 'Recalled';

        return raw;
    }

    function normalizeApprovalStepType(step) {
        var status = String((step && (step.status || step.Status)) || '').trim().toUpperCase();
        var action = String((step && (step.action || step.Action)) || '').trim().toUpperCase();

        if (action === 'APPROVE' || status.indexOf('APPROVED') === 0) return 'approve';
        if (action === 'REJECT' || status.indexOf('REJECTED') === 0) return 'reject';
        if (action === 'VETO' || status.indexOf('VETOED') === 0) return 'veto';
        if (action === 'RECALL' || status.indexOf('RECALLED') === 0) return 'recall';
        if (action === 'HOLD' || status === 'HOLD' || status.indexOf('ONHOLD') === 0 || status.indexOf('PENDINGHOLD') === 0) return 'hold';

        return 'todo';
    }

    function buildStatusCell(item, i18n, boardState) {
        var wrap = document.createElement('div');
        wrap.setAttribute(
            'style',
            'display:flex;align-items:center;gap:6px;white-space:nowrap;justify-content:center;width:100%;overflow:hidden;'
        );

        var rawStatus = String(
            item.statusCode || item.StatusCode ||
            item.status || item.Status || ''
        ).trim();

        var base = getStatusBase(rawStatus);

        var totalA = asInt(item.totalApprovers ?? item.TotalApprovers ?? 0);
        var completedA = asInt(item.completedApprovers ?? item.CompletedApprovers ?? 0);
        var actorName = String(item.resultSummary || item.ResultSummary || '').trim();

        var coopTotal = asInt(item.coopTotalSteps ?? item.CoopTotalSteps ?? 0);
        var coopDoneKeys = String(item.coopDoneKeys ?? item.CoopDoneKeys ?? '');
        var coopRejectedKeys = String(item.coopRejectedKeys ?? item.CoopRejectedKeys ?? '');
        var coopHoldKeys = String(item.coopHoldKeys ?? item.CoopHoldKeys ?? '');
        var coopRecalledKeys = String(item.coopRecalledKeys ?? item.CoopRecalledKeys ?? '');
        var coopName = String(item.coopPendingName ?? item.CoopPendingName ?? '').trim();
        var coopPos = String(item.coopPendingPosition ?? item.CoopPendingPosition ?? '').trim();

        var DOT_SLOT_WIDTH = 62;
        var TEXT_SLOT_WIDTH = 118;

        function isCompletedBoardView() {
            try {
                if (!boardState || !boardState.tab) return false;
                if (boardState.tab === 'created') return String(boardState.createdSub || '') === 'completed';
                if (boardState.tab === 'approval') return String(boardState.approvalSub || '') === 'completed';
                if (boardState.tab === 'cooperation') return String(boardState.cooperationSub || '') === 'completed';
                if (boardState.tab === 'shared') return String(boardState.sharedSub || '') === 'completed';
            } catch { }

            return false;
        }

        function parseKeySet(raw) {
            var set = new Set();

            String(raw || '')
                .split(',')
                .map(function (x) { return asInt(x); })
                .filter(function (x) { return x > 0; })
                .forEach(function (x) { set.add(x); });

            return set;
        }

        function makeDotWrap() {
            var stepWrap = document.createElement('span');
            stepWrap.setAttribute(
                'style',
                'display:inline-flex;gap:3px;align-items:center;flex-shrink:0;'
            );
            return stepWrap;
        }

        function makeLane(dotsEl, textEl) {
            var lane = document.createElement('span');
            lane.setAttribute(
                'style',
                'display:inline-grid;grid-template-columns:' + DOT_SLOT_WIDTH + 'px ' + TEXT_SLOT_WIDTH + 'px;' +
                'column-gap:6px;align-items:center;justify-content:start;flex-shrink:0;'
            );

            var dotSlot = document.createElement('span');
            dotSlot.setAttribute(
                'style',
                'width:' + DOT_SLOT_WIDTH + 'px;display:inline-flex;align-items:center;justify-content:flex-start;overflow:hidden;'
            );
            if (dotsEl) dotSlot.appendChild(dotsEl);

            var textSlot = document.createElement('span');
            textSlot.setAttribute(
                'style',
                'width:' + TEXT_SLOT_WIDTH + 'px;display:inline-flex;align-items:center;justify-content:flex-start;' +
                'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'
            );
            if (textEl) textSlot.appendChild(textEl);

            lane.appendChild(dotSlot);
            lane.appendChild(textSlot);

            return lane;
        }

        function buildProgressDots(total, done, currentType) {
            var safeTotal = asInt(total);
            var safeDone = asInt(done);
            var stepWrap = makeDotWrap();

            for (var i = 1; i <= safeTotal; i++) {
                if (i <= safeDone) stepWrap.appendChild(makeDot('approve'));
                else if (i === safeDone + 1) stepWrap.appendChild(makeDot(currentType || 'todo'));
                else stepWrap.appendChild(makeDot('todo'));
            }

            return stepWrap;
        }

        function readApprovalLaneState() {
            var steps = Array.isArray(item.approvalSteps)
                ? item.approvalSteps.slice()
                : (Array.isArray(item.ApprovalSteps) ? item.ApprovalSteps.slice() : []);

            var stateObj = {
                approved: false,
                rejected: false,
                vetoed: false,
                recalled: false,
                onHold: false,
                pending: false,
                doneCount: 0,
                steps: steps
            };

            if (steps.length > 0) {
                steps.sort(function (a, b) {
                    return asInt(a.stepOrder ?? a.StepOrder) - asInt(b.stepOrder ?? b.StepOrder);
                });

                var allApproved = steps.length > 0;

                for (var i = 0; i < steps.length; i++) {
                    var t = normalizeApprovalStepType(steps[i]);

                    if (t === 'approve') stateObj.doneCount += 1;
                    else allApproved = false;

                    if (t === 'reject') stateObj.rejected = true;
                    else if (t === 'veto') stateObj.vetoed = true;
                    else if (t === 'recall') stateObj.recalled = true;
                    else if (t === 'hold') stateObj.onHold = true;
                    else if (t === 'todo') stateObj.pending = true;
                }

                stateObj.approved =
                    allApproved &&
                    !stateObj.rejected &&
                    !stateObj.vetoed &&
                    !stateObj.recalled &&
                    !stateObj.onHold;

                if (!stateObj.approved && !stateObj.rejected && !stateObj.vetoed && !stateObj.recalled && !stateObj.onHold) {
                    stateObj.pending = true;
                }

                return stateObj;
            }

            stateObj.doneCount = completedA;

            if (base === 'Recalled') {
                stateObj.recalled = true;
                return stateObj;
            }

            if (base === 'OnHold') {
                stateObj.onHold = true;
                return stateObj;
            }

            if (base === 'Approved') {
                stateObj.approved = totalA <= 0 ? true : completedA >= totalA;
                stateObj.pending = !stateObj.approved;
                return stateObj;
            }

            if (base === 'Rejected') {
                if (totalA > 0 && completedA >= totalA) stateObj.approved = true;
                else stateObj.rejected = true;
                return stateObj;
            }

            if (base === 'Vetoed') {
                stateObj.vetoed = true;
                return stateObj;
            }

            stateObj.pending = true;
            return stateObj;
        }

        function readCoopLaneState() {
            var doneSet = parseKeySet(coopDoneKeys);
            var rejectedSet = parseKeySet(coopRejectedKeys);
            var holdSet = parseKeySet(coopHoldKeys);
            var recalledSet = parseKeySet(coopRecalledKeys);

            var stateObj = {
                approved: false,
                rejected: rejectedSet.size > 0,
                recalled: recalledSet.size > 0,
                onHold: holdSet.size > 0,
                pending: false,
                doneCount: doneSet.size,
                doneSet: doneSet,
                rejectedSet: rejectedSet,
                holdSet: holdSet,
                recalledSet: recalledSet
            };

            if (coopTotal <= 0) return stateObj;

            stateObj.approved =
                !stateObj.rejected &&
                !stateObj.recalled &&
                !stateObj.onHold &&
                doneSet.size >= coopTotal;

            stateObj.pending =
                !stateObj.approved &&
                !stateObj.rejected &&
                !stateObj.recalled &&
                !stateObj.onHold;

            return stateObj;
        }

        function resolveCompletedLaneLabelKey(laneState, otherLaneState, isDocumentRecalled) {
            if (isDocumentRecalled || laneState.recalled) return 'RecalledDone';
            if (laneState.rejected) return 'RejectedDone';
            if (laneState.vetoed) return 'VetoedDone';
            if (otherLaneState.rejected) return 'VetoedDone';
            if (laneState.approved) return 'ApprovedDone';
            if (laneState.onHold || laneState.pending) return 'Pending';
            return '';
        }

        function resolveOngoingLaneLabelKey(laneState) {
            if (laneState.recalled) return 'RecalledDone';
            if (laneState.rejected) return 'RejectedDone';
            if (laneState.vetoed) return 'VetoedDone';
            if (laneState.approved) return 'ApprovedDone';
            return '';
        }

        function buildApprovalDots(approvalState) {
            if (totalA <= 0) return null;

            var steps = approvalState.steps || [];

            if (steps.length > 0) {
                var stepWrap = makeDotWrap();

                for (var i = 0; i < steps.length; i++) {
                    stepWrap.appendChild(makeDot(normalizeApprovalStepType(steps[i])));
                }

                for (var j = steps.length + 1; j <= totalA; j++) {
                    stepWrap.appendChild(makeDot('todo'));
                }

                return stepWrap;
            }

            if (approvalState.recalled) return buildProgressDots(totalA, 0, 'recall');
            if (approvalState.rejected) return buildProgressDots(totalA, completedA, 'reject');
            if (approvalState.vetoed) return buildProgressDots(totalA, completedA, 'veto');
            if (approvalState.onHold) return buildProgressDots(totalA, completedA, 'hold');
            if (approvalState.approved) return buildProgressDots(totalA, totalA, 'approve');

            return buildProgressDots(totalA, completedA, 'todo');
        }

        function buildCoopDots(coopState) {
            if (coopTotal <= 0) return null;

            var coopWrap = makeDotWrap();

            for (var c = 1; c <= coopTotal; c++) {
                if (coopState.recalledSet.has(c)) coopWrap.appendChild(makeDot('recall'));
                else if (coopState.holdSet.has(c)) coopWrap.appendChild(makeDot('hold'));
                else if (coopState.rejectedSet.has(c)) coopWrap.appendChild(makeDot('reject'));
                else if (coopState.doneSet.has(c)) coopWrap.appendChild(makeDot('approve'));
                else coopWrap.appendChild(makeDot('todo'));
            }

            return coopWrap;
        }

        var approvalState = readApprovalLaneState();
        var coopState = readCoopLaneState();
        var docRecalled = base === 'Recalled';
        var completedView = isCompletedBoardView();

        var approvalLabelKey = completedView
            ? resolveCompletedLaneLabelKey(approvalState, coopState, docRecalled)
            : resolveOngoingLaneLabelKey(approvalState);

        var coopLabelKey = completedView
            ? resolveCompletedLaneLabelKey(coopState, approvalState, docRecalled)
            : resolveOngoingLaneLabelKey(coopState);

        var approvalTextEl = null;

        if (approvalLabelKey) {
            approvalTextEl = makeLabelSpan(getDotLabel(approvalLabelKey, i18n));
        } else if (!completedView && actorName) {
            approvalTextEl = makeNameSpan(actorName, '');
        }

        var approvalDotsEl = buildApprovalDots(approvalState);

        if (approvalDotsEl || approvalTextEl) {
            wrap.appendChild(makeLane(approvalDotsEl, approvalTextEl));
        }

        if (coopTotal > 0) {
            wrap.appendChild(makeDivider());

            var coopTextEl = null;

            if (coopLabelKey) {
                coopTextEl = makeLabelSpan(getDotLabel(coopLabelKey, i18n));
            } else if (!completedView && coopName) {
                coopTextEl = makeNameSpan(coopName, coopPos);
            }

            wrap.appendChild(makeLane(buildCoopDots(coopState), coopTextEl));
        }

        return wrap;
    }

    function readCurrentUserId(root) {
        var dataCurrentUserId = '';

        try {
            var anyEl = document.querySelector('[data-current-user-id]');
            dataCurrentUserId = anyEl ? (anyEl.getAttribute('data-current-user-id') || '') : '';
        } catch { }

        var candidates = [
            root && root.dataset ? root.dataset.currentUserId : '',
            root && root.dataset ? root.dataset.currentUserid : '',
            root && root.dataset ? root.dataset.userId : '',
            root && root.dataset ? root.dataset.userid : '',
            document.body && document.body.dataset ? document.body.dataset.currentUserId : '',
            document.body && document.body.dataset ? document.body.dataset.currentUserid : '',
            document.body && document.body.dataset ? document.body.dataset.userId : '',
            document.body && document.body.dataset ? document.body.dataset.userid : '',
            (document.querySelector('meta[name="current-user-id"]') || {}).content || '',
            (document.querySelector('meta[name="CurrentUserId"]') || {}).content || '',
            (document.querySelector('input[name="CurrentUserId"]') || {}).value || '',
            (document.querySelector('input[name="currentUserId"]') || {}).value || '',
            dataCurrentUserId,
            window.__CURRENT_USER_ID__ || '',
            window.currentUserId || ''
        ];

        for (var i = 0; i < candidates.length; i++) {
            var id = normalizeId(candidates[i]);
            if (id) return id;
        }

        return '';
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

    function createBoard(root) {
        if (!root || !root.dataset) {
            console.error('[DocBoard] root or dataset missing');
            return;
        }

        if (!window.DevExpress || !DevExpress.ui || !DevExpress.ui.dxDataGrid || !DevExpress.data || !DevExpress.data.CustomStore) {
            console.error('[DocBoard] DevExtreme DataGrid가 로드되지 않았습니다.');
            return;
        }

        __docBoardCurrentUserId = readCurrentUserId(root);

        var apiList = String(root.dataset.apiList || '');
        var apiBadges = String(root.dataset.apiBadges || '');
        var detailUrl = String(root.dataset.detailUrl || '');

        if (!apiList) {
            console.error('[DocBoard] data-api-list is empty. root:', root);
            return;
        }

        var gridEl = document.getElementById('docBoardGrid');
        if (!gridEl) {
            console.error('[DocBoard] #docBoardGrid not found.');
            return;
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

        var bulkTools = document.getElementById('docBulkTools');
        var chkAllBottom = document.getElementById('chkAllBottom');
        var btnApproveBulk = document.getElementById('btnApproveBulk');

        var i18n = safeParseI18n();
        var H = window.DocControllerHelper;

        if (!H) {
            console.error('[DocBoard] DocControllerHelper.js가 로드되지 않았습니다.');
            gridEl.innerHTML = '<div class="alert alert-danger m-3">' + String((i18n && i18n.BULK_FAIL) || 'Error') + '</div>';
            return;
        }

        var boardDateText = {
            currentCulture: String((root.dataset && root.dataset.currentCulture) || i18n.CURRENT_CULTURE || 'ko-KR'),
            datePattern: String((root.dataset && root.dataset.datePattern) || i18n.DATE_PATTERN || H.getLocalDatePattern((root.dataset && root.dataset.currentCulture) || i18n.CURRENT_CULTURE || 'ko-KR')),
            ok: String(i18n.DATE_OK || '확인'),
            cancel: String(i18n.DATE_CANCEL || '취소'),
            today: String(i18n.DATE_TODAY || '금일'),
            empty: ''
        };

        var READ_LABEL = String(i18n.READ_LABEL || 'Viewed');
        var UNREAD_LABEL = String(i18n.UNREAD_LABEL || 'Unviewed');
        var LIST_LOADING = String(i18n.LIST_LOADING || 'Loading...');
        var LIST_EMPTY = String(i18n.LIST_EMPTY || 'Empty');
        var BULK_OK_FMT = String(i18n.DOC_Toast_BulkApproved || '');
        var BULK_FAIL = String(i18n.BULK_FAIL || '');
        var GRID_ALL = String(i18n.GRID_ALL || 'All');
        var BULK_APPROVE = String(i18n.BULK_APPROVE || '승인');

        var selectedDocIds = new Set();
        var lastTotal = 0;
        var grid = null;
        var cachedStatusHeaderItems = [];

        var allowedTabs = ['created', 'approval', 'cooperation', 'shared'];
        var STORAGE_KEY = 'docBoardGridStateByTab';

        var state = {
            tab: 'created',
            page: 1,
            pageSize: 20,
            createdSub: 'ongoing',
            approvalSub: 'ongoing',
            cooperationSub: 'ongoing',
            sharedSub: 'ongoing'
        };

        function isApprovalOngoing() {
            return state.tab === 'approval' && String(state.approvalSub || '') === 'ongoing';
        }

        function readStateMap() {
            try {
                var s = window.sessionStorage.getItem(STORAGE_KEY);
                if (!s) return {};
                var obj = JSON.parse(s);
                return (obj && typeof obj === 'object') ? obj : {};
            } catch {
                return {};
            }
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
                createdSub: state.createdSub,
                approvalSub: state.approvalSub,
                cooperationSub: state.cooperationSub,
                sharedSub: state.sharedSub
            };

            writeStateMap(map);
        }

        function readStateForTabFromSession(tab) {
            var map = readStateMap();
            var s = map && map[tab];

            return (s && typeof s === 'object') ? s : null;
        }

        function setUrlFromState(replaceOnly) {
            try {
                var qs = new URLSearchParams(window.location.search || '');

                qs.set('tab', state.tab);
                qs.set('page', String(state.page || 1));
                qs.set('pageSize', String(state.pageSize || 20));

                if (state.tab === 'created') qs.set('createdSub', String(state.createdSub || 'ongoing')); else qs.delete('createdSub');
                if (state.tab === 'approval') qs.set('approvalSub', String(state.approvalSub || 'ongoing')); else qs.delete('approvalSub');
                if (state.tab === 'cooperation') qs.set('cooperationSub', String(state.cooperationSub || 'ongoing')); else qs.delete('cooperationSub');
                if (state.tab === 'shared') qs.set('sharedSub', String(state.sharedSub || 'ongoing')); else qs.delete('sharedSub');

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

            var cs = qs.get('createdSub'); if (cs) out.createdSub = cs;
            var as = qs.get('approvalSub'); if (as) out.approvalSub = as;
            var cos = qs.get('cooperationSub'); if (cos) out.cooperationSub = cos;
            var ss = qs.get('sharedSub'); if (ss) out.sharedSub = ss;

            return out;
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
            tabs.forEach(function (x) {
                x.setAttribute('aria-selected', x.dataset.tab === state.tab ? 'true' : 'false');
            });

            applySubtabsUi();
            updateBulkUiState();
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
            state.createdSub = fromUrl.createdSub || fromSession.createdSub || 'ongoing';
            state.approvalSub = fromUrl.approvalSub || fromSession.approvalSub || 'ongoing';
            state.cooperationSub = fromUrl.cooperationSub || fromSession.cooperationSub || 'ongoing';
            state.sharedSub = fromUrl.sharedSub || fromSession.sharedSub || 'ongoing';

            try { window.sessionStorage.setItem('docBoardTab', state.tab); } catch { }

            applyTabUi();
            setUrlFromState(true);
            saveStateToSession();
        }

        function setBadgeDom(id, val) {
            var el = document.getElementById(id);
            if (!el) return;

            var n = asInt(val);
            el.textContent = String(n);
            el.hidden = !(n > 0);
        }

        async function refreshBadges() {
            if (!apiBadges) return;

            try {
                var url = apiBadges + (apiBadges.indexOf('?') >= 0 ? '&' : '?') + '_ts=' + Date.now();
                var res = await fetch(url, {
                    method: 'GET',
                    cache: 'no-store',
                    headers: { 'Accept': 'application/json' }
                });

                var j = await res.json().catch(function () { return null; });
                if (!res.ok || !j) return;

                setBadgeDom('badge-created', j.createdUnread ?? j.CreatedUnread ?? j.created ?? j.Created ?? 0);
                setBadgeDom('badge-approval', j.approvalPending ?? j.ApprovalPending ?? j.approval ?? j.Approval ?? 0);
                setBadgeDom('badge-cooperation', j.cooperationPending ?? j.CooperationPending ?? j.cooperation ?? j.Cooperation ?? 0);
                setBadgeDom('badge-shared', j.sharedUnread ?? j.SharedUnread ?? j.shared ?? j.Shared ?? 0);
            } catch { }
        }

        function getReadText(item) {
            return toBool(pick(item, ['isRead', 'IsRead', 'read', 'Read', 'is_viewed', 'isViewed', 'IsViewed']))
                ? READ_LABEL
                : UNREAD_LABEL;
        }

        function getTerminalStatusLabel(item) {
            var status = norm(item.statusCode || item.StatusCode || item.status || item.Status);

            if (status === 'APPROVED' || status === 'APPROVE') return getDotLabel('ApprovedDone', i18n);
            if (status === 'REJECTED' || status === 'REJECT') return getDotLabel('RejectedDone', i18n);
            if (status === 'VETOED' || status === 'VETO') return getDotLabel('VetoedDone', i18n);
            if (status === 'RECALLED' || status === 'RECALL') return getDotLabel('RecalledDone', i18n);
            if (status === 'HOLD' || status === 'ONHOLD' || status === 'PENDINGHOLD') return String((i18n.STATUS_LABELS || {}).OnHold || '보류');

            var approvalSteps = Array.isArray(item.approvalSteps)
                ? item.approvalSteps
                : (Array.isArray(item.ApprovalSteps) ? item.ApprovalSteps : []);

            if (approvalSteps.some(function (x) { return normalizeApprovalStepType(x) === 'reject'; })) return getDotLabel('RejectedDone', i18n);
            if (approvalSteps.some(function (x) { return normalizeApprovalStepType(x) === 'veto'; })) return getDotLabel('VetoedDone', i18n);
            if (approvalSteps.some(function (x) { return normalizeApprovalStepType(x) === 'recall'; })) return getDotLabel('RecalledDone', i18n);
            if (approvalSteps.some(function (x) { return normalizeApprovalStepType(x) === 'hold'; })) return String((i18n.STATUS_LABELS || {}).OnHold || '보류');

            if (approvalSteps.length > 0 && approvalSteps.every(function (x) { return normalizeApprovalStepType(x) === 'approve'; })) {
                return getDotLabel('ApprovedDone', i18n);
            }

            return '';
        }

        function countKeyList(value) {
            return String(value || '')
                .split(',')
                .map(function (x) { return asInt(x); })
                .filter(function (x) { return x > 0; })
                .length;
        }

        function hasAnyKeyList(value) {
            return countKeyList(value) > 0;
        }

        function getApprovalVisibleText(item) {
            var terminalLabel = getTerminalStatusLabel(item);
            if (terminalLabel) return terminalLabel;

            return String(item.resultSummary || item.ResultSummary || '').trim();
        }

        function getCoopVisibleText(item) {
            var total = asInt(item.coopTotalSteps ?? item.CoopTotalSteps);
            if (total <= 0) return '';

            var terminalLabel = getTerminalStatusLabel(item);
            if (terminalLabel) return terminalLabel;

            if (hasAnyKeyList(item.coopRejectedKeys ?? item.CoopRejectedKeys)) return getDotLabel('RejectedDone', i18n);
            if (hasAnyKeyList(item.coopRecalledKeys ?? item.CoopRecalledKeys)) return getDotLabel('RecalledDone', i18n);
            if (hasAnyKeyList(item.coopHoldKeys ?? item.CoopHoldKeys)) return String((i18n.STATUS_LABELS || {}).OnHold || '보류');

            if (countKeyList(item.coopDoneKeys ?? item.CoopDoneKeys) >= total) {
                return getDotLabel('ApprovedDone', i18n);
            }

            var name = String(item.coopPendingName || item.CoopPendingName || '').trim();
            var pos = String(item.coopPendingPosition || item.CoopPendingPosition || '').trim();

            return [name, pos].filter(function (x) { return !!x; }).join(' ');
        }

        function buildStatusSearchText(item) {
            var parts = [];

            addUnique(parts, getApprovalVisibleText(item));
            addUnique(parts, getCoopVisibleText(item));
            addUnique(parts, item.statusCode || item.StatusCode);
            addUnique(parts, item.status || item.Status);

            return parts.join(' ');
        }

        function buildStatusFilterItems(item) {
            var items = [];

            function push(key, text) {
                key = String(key || '').trim();
                text = String(text || '').trim();
                if (!key || !text) return;
                if (items.some(function (x) { return x.key === key; })) return;
                items.push({ key: key, text: text });
            }

            var terminalLabel = getTerminalStatusLabel(item);
            var status = norm(item.statusCode || item.StatusCode || item.status || item.Status);

            if (terminalLabel) {
                if (status === 'APPROVED' || status === 'APPROVE') push('STATUS_APPROVED', getDotLabel('ApprovedDone', i18n));
                else if (status === 'REJECTED' || status === 'REJECT') push('STATUS_REJECTED', getDotLabel('RejectedDone', i18n));
                else if (status === 'VETOED' || status === 'VETO') push('STATUS_VETOED', getDotLabel('VetoedDone', i18n));
                else if (status === 'RECALLED' || status === 'RECALL') push('STATUS_RECALLED', getDotLabel('RecalledDone', i18n));
                else if (status === 'HOLD' || status === 'ONHOLD' || status === 'PENDINGHOLD') push('STATUS_ONHOLD', String((i18n.STATUS_LABELS || {}).OnHold || '보류'));
                else push('STATUS_PENDING', terminalLabel);
                return items;
            }

            var approvalName = getApprovalVisibleText(item);
            if (approvalName) {
                push('APPROVAL_PENDING_TEXT:' + approvalName, approvalName);
            }

            var coopName = getCoopVisibleText(item);
            if (coopName) {
                push('COOP_PENDING_TEXT:' + coopName, coopName);
            }

            return items;
        }

        function buildStatusFilterKeyText(item) {
            var items = buildStatusFilterItems(item);
            var keys = [];

            items.forEach(function (x) {
                if (x.key && keys.indexOf(x.key) < 0) keys.push(x.key);
            });

            return keys.length > 0 ? '|' + keys.join('|') + '|' : '';
        }

        function mergeStatusHeaderItems(items) {
            var base = [
                { text: getDotLabel('ApprovedDone', i18n), key: 'STATUS_APPROVED' },
                { text: getDotLabel('RejectedDone', i18n), key: 'STATUS_REJECTED' },
                { text: getDotLabel('VetoedDone', i18n), key: 'STATUS_VETOED' },
                { text: getDotLabel('RecalledDone', i18n), key: 'STATUS_RECALLED' },
                { text: String((i18n.STATUS_LABELS || {}).OnHold || '보류'), key: 'STATUS_ONHOLD' }
            ];

            var map = {};
            var merged = [];

            function add(item) {
                if (!item) return;

                var key = String(item.key || item.value || '').trim();
                var text = String(item.text || '').trim();

                if (!key || !text) return;
                if (map[key]) return;

                map[key] = true;

                merged.push({
                    text: text,
                    value: ['statusFilterKeyText', 'contains', '|' + key + '|']
                });
            }

            base.forEach(add);

            (items || []).forEach(function (item) {
                add(item);
            });

            return merged;
        }

        function prepareItems(res) {
            var total = asInt(res && res.total);
            var page = asInt(res && res.page);
            var pageSize = asInt(res && res.pageSize);
            var startIndex = (page - 1) * pageSize;
            var dynamicStatusItems = [];

            lastTotal = total;

            var items = ((res && res.items) || []).map(function (item, idx) {
                var row = item || {};
                var displayNo = total - startIndex - idx;

                row.CreatedAtLocalText = H.getFieldValue(row, ['CreatedAtLocalText', 'createdAtLocalText'], '');
                row.createdAtLocalText = row.CreatedAtLocalText;
                row.CreatedAtLocalDateKey = H.getFieldValue(row, ['CreatedAtLocalDateKey', 'createdAtLocalDateKey'], '');
                row.createdAtLocalDateKey = row.CreatedAtLocalDateKey;
                row.CreatedAtUtc = H.getFieldValue(row, ['CreatedAtUtc', 'createdAtUtc'], '');
                row.createdAtUtc = row.CreatedAtUtc;
                row.CreatedAt = row.CreatedAtLocalText;
                row.createdAt = row.CreatedAtLocalText;

                H.normalizeLocalDateRow(row, {
                    localTextField: 'CreatedAtLocalText',
                    localDateKeyField: 'CreatedAtLocalDateKey'
                });

                row.__displayNo = displayNo;
                row.readText = getReadText(row);
                row.statusSearchText = buildStatusSearchText(row);
                row.statusFilterKeyText = buildStatusFilterKeyText(row);

                buildStatusFilterItems(row).forEach(function (x) {
                    if (!x || !x.key || !x.text) return;

                    dynamicStatusItems.push({
                        text: x.text,
                        key: x.key
                    });
                });

                return row;
            });

            cachedStatusHeaderItems = mergeStatusHeaderItems(dynamicStatusItems);

            return items;
        }

        function statusHeaderFilterDataSource(options) {
            var data = cachedStatusHeaderItems && cachedStatusHeaderItems.length
                ? cachedStatusHeaderItems
                : mergeStatusHeaderItems([]);

            if (options) {
                options.dataSource = data;
            }

            return data;
        }

        function mapDxSortToLegacy(sort) {
            if (!sort || !sort.length) return 'created_desc';

            var first = sort[0] || {};
            var selector = String(first.selector || '');
            var selectorKey = selector.toLowerCase();
            var desc = first.desc === true;

            if (selectorKey === 'createdat' || selectorKey === 'createdatlocaltext' || selectorKey === 'createdatlocaldatekey') return desc ? 'created_desc' : 'created_asc';
            if (selectorKey === 'templatetitle') return desc ? 'title_desc' : 'title_asc';

            return 'created_desc';
        }

        function buildBoardDataUrl(loadOptions) {
            var take = asInt(loadOptions && loadOptions.take);
            var skip = asInt(loadOptions && loadOptions.skip);

            if (take <= 0) take = state.pageSize || 20;

            var page = Math.floor(skip / take) + 1;
            if (page < 1) page = 1;

            state.page = page;
            state.pageSize = take;

            var params = new URLSearchParams();

            params.set('tab', state.tab || 'created');
            params.set('page', String(page));
            params.set('pageSize', String(take));
            params.set('titleFilter', 'all');
            params.set('sort', mapDxSortToLegacy(loadOptions && loadOptions.sort));
            params.set('q', '');

            var requestCulture = String(boardDateText.currentCulture || i18n.CURRENT_CULTURE || '').trim();
            if (requestCulture) {
                params.set('culture', requestCulture);
                params.set('ui-culture', requestCulture);
            }

            if (state.tab === 'created') params.set('createdSub', String(state.createdSub || 'ongoing'));
            if (state.tab === 'approval') params.set('approvalSub', String(state.approvalSub || 'ongoing'));
            if (state.tab === 'cooperation') params.set('cooperationSub', String(state.cooperationSub || 'ongoing'));
            if (state.tab === 'shared') params.set('sharedSub', String(state.sharedSub || 'ongoing'));

            if (loadOptions && loadOptions.filter) {
                params.set('dxFilter', JSON.stringify(loadOptions.filter));
            }

            if (loadOptions && loadOptions.sort) {
                params.set('dxSort', JSON.stringify(loadOptions.sort));
            }

            params.set('_ts', String(Date.now()));

            setUrlFromState(true);
            saveStateToSession();

            return apiList + (apiList.indexOf('?') >= 0 ? '&' : '?') + params.toString();
        }

        function updateBulkUiState() {
            var show = isApprovalOngoing();

            if (bulkTools) {
                if (show) bulkTools.classList.add('is-visible');
                else bulkTools.classList.remove('is-visible');
            }

            if (!show) {
                selectedDocIds.clear();

                if (chkAllBottom) {
                    chkAllBottom.checked = false;
                }

                if (btnApproveBulk) {
                    btnApproveBulk.disabled = true;
                }

                return;
            }

            var visibleRows = grid
                ? grid.getVisibleRows().filter(function (x) { return x && x.rowType === 'data' && x.data; })
                : [];

            var total = visibleRows.length;
            var checked = 0;

            visibleRows.forEach(function (r) {
                var id = String((r.data && (r.data.docId || r.data.DocId)) || '').trim();
                if (id && selectedDocIds.has(id)) checked++;
            });

            if (chkAllBottom) {
                chkAllBottom.checked = total > 0 && checked === total;
            }

            if (btnApproveBulk) {
                btnApproveBulk.textContent = BULK_APPROVE;
                btnApproveBulk.disabled = selectedDocIds.size === 0;
            }
        }

        function setAllVisibleChecks(checked) {
            if (!grid || !isApprovalOngoing()) return;

            var rows = grid.getVisibleRows().filter(function (x) {
                return x && x.rowType === 'data' && x.data;
            });

            rows.forEach(function (r) {
                var id = String((r.data.docId || r.data.DocId || '')).trim();
                if (!id) return;

                if (checked) selectedDocIds.add(id);
                else selectedDocIds.delete(id);
            });

            grid.repaint();
            updateBulkUiState();
        }

        async function doBulkApprove(docIds) {
            var token = getCsrfToken();

            var res = await fetch('/Doc/BulkApprove', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'RequestVerificationToken': token
                },
                credentials: 'same-origin',
                body: JSON.stringify({ DocIds: docIds || [] })
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

        async function handleBulkApprove() {
            if (!isApprovalOngoing()) return;

            var ids = Array.from(selectedDocIds.values())
                .map(function (x) { return String(x || '').trim(); })
                .filter(function (x) { return !!x; });

            if (ids.length === 0) return;

            if (btnApproveBulk) btnApproveBulk.disabled = true;

            try {
                var j = await doBulkApprove(ids);

                selectedDocIds.clear();

                if (chkAllBottom) chkAllBottom.checked = false;

                await refreshBadges();

                if (grid) {
                    await grid.getDataSource().reload();
                }

                var approved = 0;

                try {
                    approved = asInt(j && j.totals ? (j.totals.approved ?? j.totals.Approved) : 0);
                } catch {
                    approved = 0;
                }

                if (BULK_OK_FMT) {
                    showTopBarToast(fmt0(BULK_OK_FMT, approved), 'success');
                }
            } catch (e) {
                console.error('[DocBoard] BulkApprove failed:', e && e.payload ? e.payload : e);

                if (BULK_FAIL) {
                    showTopBarToast(BULK_FAIL, 'danger');
                }
            } finally {
                updateBulkUiState();
            }
        }

        function reloadGrid(resetPage) {
            selectedDocIds.clear();

            if (chkAllBottom) {
                chkAllBottom.checked = false;
            }

            if (!grid) return;

            if (resetPage) {
                var currentPageIndex = 0;

                try {
                    currentPageIndex = grid.pageIndex();
                } catch {
                    currentPageIndex = 0;
                }

                if (currentPageIndex !== 0) {
                    grid.pageIndex(0);
                    updateBulkUiState();
                    return;
                }
            }

            grid.getDataSource().reload();
            updateBulkUiState();
        }

        function createGrid() {
            var store = new DevExpress.data.CustomStore({
                key: 'docId',
                load: function (loadOptions) {
                    var url = buildBoardDataUrl(loadOptions || {});

                    return fetch(url, {
                        method: 'GET',
                        cache: 'no-store',
                        headers: {
                            'Accept': 'application/json'
                        }
                    })
                        .then(function (r) {
                            return r.json().then(function (j) {
                                if (!r.ok || !j) {
                                    throw new Error('BoardData load failed');
                                }

                                var items = prepareItems(j);

                                return {
                                    data: items,
                                    totalCount: asInt(j.total)
                                };
                            });
                        });
                }
            });

            grid = new DevExpress.ui.dxDataGrid(gridEl, {
                dataSource: store,
                showBorders: true,
                rowAlternationEnabled: true,
                hoverStateEnabled: true,
                allowColumnResizing: true,
                columnResizingMode: 'widget',
                columnAutoWidth: false,
                wordWrapEnabled: false,
                noDataText: LIST_EMPTY,
                remoteOperations: {
                    paging: true,
                    filtering: true,
                    sorting: true
                },
                scrolling: {
                    mode: 'standard',
                    columnRenderingMode: 'standard',
                    showScrollbar: 'never',
                    useNative: true
                },
                paging: {
                    pageSize: state.pageSize || 20,
                    pageIndex: Math.max(0, (state.page || 1) - 1)
                },
                pager: {
                    visible: true,
                    showInfo: true,
                    showNavigationButtons: true,
                    showPageSizeSelector: true,
                    allowedPageSizes: [20, 50, 100]
                },
                sorting: {
                    mode: 'multiple'
                },
                filterRow: {
                    visible: true,
                    applyFilter: 'auto',
                    showAllText: GRID_ALL
                },
                headerFilter: {
                    visible: true
                },
                searchPanel: {
                    visible: false
                },
                loadPanel: {
                    enabled: true,
                    text: LIST_LOADING
                },
                columns: [
                    {
                        caption: String(i18n.COL_NO || '번호'),
                        dataField: '__displayNo',
                        width: 86,
                        alignment: 'center',
                        allowFiltering: false,
                        allowSorting: false,
                        allowHeaderFiltering: false,
                        headerCellTemplate: function (container) {
                            var el = container && container.jquery ? container[0] : container;
                            var host = document.createElement('span');
                            host.className = 'doc-no-cell';

                            if (isApprovalOngoing()) {
                                var cb = document.createElement('input');
                                cb.type = 'checkbox';
                                cb.addEventListener('change', function () {
                                    setAllVisibleChecks(cb.checked);
                                });
                                host.appendChild(cb);
                            }

                            var label = document.createElement('span');
                            label.textContent = String(i18n.COL_NO || '번호');
                            host.appendChild(label);

                            el.appendChild(host);
                        },
                        cellTemplate: function (container, options) {
                            var el = container && container.jquery ? container[0] : container;
                            var host = document.createElement('span');
                            host.className = 'doc-no-cell';

                            var docId = String((options.data && (options.data.docId || options.data.DocId)) || '').trim();

                            if (isApprovalOngoing()) {
                                var cb = document.createElement('input');
                                cb.type = 'checkbox';
                                cb.setAttribute('data-doccheck', '1');
                                cb.checked = docId ? selectedDocIds.has(docId) : false;
                                cb.addEventListener('change', function () {
                                    if (!docId) return;

                                    if (cb.checked) selectedDocIds.add(docId);
                                    else selectedDocIds.delete(docId);

                                    updateBulkUiState();
                                });
                                host.appendChild(cb);
                            }

                            var noText = document.createElement('span');
                            noText.textContent = String(options.value || '');
                            host.appendChild(noText);

                            el.appendChild(host);
                        }
                    },
                    {
                        caption: String(i18n.COL_TITLE || '제목'),
                        dataField: 'templateTitle',
                        width: 320,
                        alignment: 'left',
                        allowHeaderFiltering: false,
                        cellTemplate: function (container, options) {
                            var el = container && container.jquery ? container[0] : container;
                            var row = options.data || {};
                            var docId = String(row.docId || row.DocId || '').trim();
                            var title = String(row.templateTitle || row.TemplateTitle || '').trim();

                            var wrapper = document.createElement('span');
                            wrapper.className = 'doc-title';

                            var link = document.createElement('a');
                            var tabQs = state && state.tab ? ('?tab=' + encodeURIComponent(state.tab)) : '';
                            link.href = detailUrl.replace('{docId}', encodeURIComponent(docId)) + tabQs;
                            link.title = title;
                            link.textContent = title;

                            var rawIsRead = pick(row, ['isRead', 'IsRead', 'read', 'Read', 'is_viewed', 'isViewed', 'IsViewed']);
                            var isRead = toBool(rawIsRead);

                            var isMyPendingTurn = toBool(pick(row, ['isMyPendingTurn', 'IsMyPendingTurn']));
                            var isMyPendingCooperation = toBool(pick(row, ['isMyPendingCooperation', 'IsMyPendingCooperation']));

                            var shouldBold =
                                (state.tab === 'approval' && String(state.approvalSub || '') === 'ongoing' && isMyPendingTurn) ||
                                (state.tab === 'cooperation' && String(state.cooperationSub || '') === 'ongoing' && isMyPendingCooperation) ||
                                ((state.tab === 'created' || state.tab === 'shared') && !isRead);

                            if (shouldBold) link.classList.add('doc-unread');

                            wrapper.appendChild(link);
                            el.appendChild(wrapper);
                        }
                    },
                    {
                        caption: String(i18n.COL_AUTHOR || '작성자'),
                        dataField: 'authorName',
                        width: 180,
                        alignment: 'left',
                        allowHeaderFiltering: false
                    },
                    H.createLocalDateColumn({
                        caption: String(i18n.COL_DATE || '작성일시'),
                        text: {
                            currentCulture: i18n.CURRENT_CULTURE || 'ko-KR',
                            datePattern: i18n.DATE_PATTERN || 'yyyy-MM-dd',
                            ok: i18n.DATE_OK || '확인',
                            cancel: i18n.DATE_CANCEL || '취소',
                            today: i18n.DATE_TODAY || '금일',
                            empty: ''
                        },
                        localTextField: 'CreatedAtLocalText',
                        localDateKeyField: 'CreatedAtLocalDateKey',
                        sortFields: [
                            'CreatedAtLocalDateKey',
                            'createdAtLocalDateKey',
                            'CreatedAtLocalText',
                            'createdAtLocalText'
                        ],
                        width: 170,
                        alignment: 'center',
                        allowHeaderFiltering: false,
                        allowFiltering: true,
                        allowSorting: true
                    }),
                    {
                        caption: String(i18n.COL_STATUS || '상태'),
                        dataField: 'statusSearchText',
                        width: 430,
                        alignment: 'center',
                        allowHeaderFiltering: true,
                        headerFilter: {
                            dataSource: statusHeaderFilterDataSource
                        },
                        cellTemplate: function (container, options) {
                            var el = container && container.jquery ? container[0] : container;
                            el.appendChild(buildStatusCell(options.data || {}, i18n, state));
                        }
                    },
                    {
                        caption: String(i18n.COL_READ || '열람'),
                        dataField: 'isRead',
                        dataType: 'boolean',
                        width: 150,
                        alignment: 'center',
                        trueText: READ_LABEL,
                        falseText: UNREAD_LABEL,
                        lookup: {
                            dataSource: [
                                { value: true, text: READ_LABEL },
                                { value: false, text: UNREAD_LABEL }
                            ],
                            valueExpr: 'value',
                            displayExpr: 'text'
                        },
                        headerFilter: {
                            dataSource: [
                                { text: READ_LABEL, value: true },
                                { text: UNREAD_LABEL, value: false }
                            ]
                        },
                        cellTemplate: function (container, options) {
                            var el = container && container.jquery ? container[0] : container;
                            el.textContent = options.value === true ? READ_LABEL : UNREAD_LABEL;
                        }
                    }
                ],
                onContentReady: function () {
                    updateBulkUiState();
                },
                onOptionChanged: function (e) {
                    if (!e) return;

                    if (e.fullName === 'paging.pageIndex') {
                        state.page = asInt(e.value) + 1;
                        setUrlFromState(true);
                        saveStateToSession();
                    }

                    if (e.fullName === 'paging.pageSize') {
                        state.pageSize = asInt(e.value) || 20;
                        state.page = 1;
                        setUrlFromState(true);
                        saveStateToSession();
                    }
                }
            });
        }

        async function forceRefreshBoard(reason) {
            try {
                selectedDocIds.clear();

                if (chkAllBottom) chkAllBottom.checked = false;

                console.debug('[DocBoard] force refresh by', reason || '');

                await refreshBadges();

                if (grid) {
                    await grid.getDataSource().reload();
                }
            } catch (e) {
                console.error('[DocBoard] force refresh failed:', e);
            } finally {
                updateBulkUiState();
            }
        }

        function isBackForwardNavigation(ev) {
            try {
                if (ev && ev.persisted) return true;
            } catch { }

            try {
                if (window.performance && typeof window.performance.getEntriesByType === 'function') {
                    var nav = window.performance.getEntriesByType('navigation');

                    if (nav && nav.length > 0 && nav[0] && nav[0].type === 'back_forward') {
                        return true;
                    }
                }
            } catch { }

            return false;
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
                    reloadGrid(true);
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
                    state.createdSub = s.createdSub || state.createdSub || 'ongoing';
                    state.approvalSub = s.approvalSub || state.approvalSub || 'ongoing';
                    state.cooperationSub = s.cooperationSub || state.cooperationSub || 'ongoing';
                    state.sharedSub = s.sharedSub || state.sharedSub || 'ongoing';
                } else {
                    state.page = 1;
                }

                selectedDocIds.clear();

                try { window.sessionStorage.setItem('docBoardTab', state.tab); } catch { }

                applyTabUi();
                setUrlFromState(true);
                saveStateToSession();
                refreshBadges();
                reloadGrid(true);
            });
        });

        if (chkAllBottom) {
            chkAllBottom.addEventListener('change', function () {
                setAllVisibleChecks(chkAllBottom.checked);
            });
        }

        if (btnApproveBulk) {
            btnApproveBulk.addEventListener('click', function () {
                handleBulkApprove();
            });
        }

        window.addEventListener('popstate', function () {
            var u = readStateFromUrl();

            if (u.tab && allowedTabs.indexOf(u.tab) >= 0) state.tab = u.tab;

            state.page = u.page || 1;
            state.pageSize = u.pageSize || 20;

            if (u.createdSub) state.createdSub = u.createdSub;
            if (u.approvalSub) state.approvalSub = u.approvalSub;
            if (u.cooperationSub) state.cooperationSub = u.cooperationSub;
            if (u.sharedSub) state.sharedSub = u.sharedSub;

            selectedDocIds.clear();

            applyTabUi();
            saveStateToSession();

            if (grid) {
                grid.option('paging.pageSize', state.pageSize);
                grid.pageIndex(Math.max(0, state.page - 1));
                grid.getDataSource().reload();
            }

            refreshBadges();
        });

        window.addEventListener('pageshow', function (ev) {
            if (!isBackForwardNavigation(ev)) return;
            forceRefreshBoard('pageshow');
        });

        document.addEventListener('visibilitychange', function () {
            if (document.visibilityState !== 'visible') return;
            if (!document.getElementById('docBoard')) return;
            forceRefreshBoard('visibilitychange');
        });

        restoreInitialState();
        createGrid();
        refreshBadges();
    }

    window.EBDocStatusDots = window.EBDocStatusDots || {};
    window.EBDocStatusDots.renderStatusCell = function (item, i18n, boardState) {
        return buildStatusCell(item || {}, i18n || {}, boardState || {});
    };

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

    function autoInitDocBoard() {
        var root = document.getElementById('docBoard') || document.querySelector('[data-api-list]');

        if (root) {
            console.info('[DocBoard] auto init. apiList =', root.dataset ? root.dataset.apiList : '');
            window.EBDocBoard.init(root);
            return;
        }

        var path = String(window.location && window.location.pathname ? window.location.pathname : '').toLowerCase();

        if (path.indexOf('/doc/board') === 0) {
            console.warn('[DocBoard] root not found for auto init');
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', autoInitDocBoard, { once: true });
    } else {
        autoInitDocBoard();
    }
})();