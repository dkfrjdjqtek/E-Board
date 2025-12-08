
// 2025.11.14 Changed: Excel 폰트 무시하고 eb.doc.preview.css 기준 폰트 통일(applyStyleToCell 폰트 적용 제거)
// 2025.11.14 Changed: colorizeInputGroups 폴백 추가 및 mount 안전화 프리뷰 실패 방지
// 2025.11.14 Changed: Compose Detail 폰트 px 고정과 동일 줄바꿈 일치 위해 측정폭 계산에 border 포함 cellc lineHeight=rowPx 강제 copyStyle 자간 normal 고정 hideRingIfIdle 안정화
// 2025.11.14 Added: DOC Compose Detail 공통 프리뷰 렌더링 스크립트 분리
(function () {
    // 전역 폴백 colorizeInputGroups 미정의 시 no-op 제공
    if (typeof window.colorizeInputGroups !== 'function') {
        window.colorizeInputGroups = function () { /* no-op */ };
    }

    try {
        var els = document.querySelectorAll('.doc-unclip');
        for (var i = 0; i < els.length; i++) {
            els[i].style.setProperty('contain', 'none', 'important');
        }
    } catch (e) { /* no-op */ }

    // 쓰기모드 플래그
    let __EB_IS_WRITE__ = (typeof window.__DOC_IS_WRITE__ === 'boolean') ? !!window.__DOC_IS_WRITE__ : true;
    window.__DOC_IS_WRITE__ = __EB_IS_WRITE__;

    /* ===== 공통 유틸 ===== */
    const $ = (q, r = document) => r.querySelector(q);
    const $$ = (q, r = document) => Array.from(r.querySelectorAll(q));
    const $alert = () => document.getElementById('doc-alert');

    function decodeHtmlEntities(str) {
        if (str == null) return '';
        let s = String(str);
        s = s.replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)));
        s = s.replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)));
        s = s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
        return s;
    }

    const T = (m) => {
        if (m == null) return '';
        if (Array.isArray(m)) return m.map(T).join('\n');
        const s = String(m);
        const mapped = (window.__RESX && Object.prototype.hasOwnProperty.call(window.__RESX, s)) ? window.__RESX[s] : s;
        return decodeHtmlEntities(mapped);
    };

    const info = m => { $alert()?.replaceChildren(Object.assign(document.createElement('div'), { className: 'alert alert-info', textContent: T(m) })); };
    const ok = m => { $alert()?.replaceChildren(Object.assign(document.createElement('div'), { className: 'alert alert-success', textContent: T(m) })); };
    const err = m => { $alert()?.replaceChildren(Object.assign(document.createElement('div'), { className: 'alert alert-danger', textContent: T(m || 'DOC_Err_PreviewFailed') })); };

    /* ===== 레이아웃 고정 ===== */
    function escapeFixedTraps(el, maxHops = 10, maxApply = 4) {
        let n = el.parentElement, h = 0, a = 0;
        while (n && h < maxHops && a < maxApply) {
            const cs = getComputedStyle(n);
            const hasTrap = (cs.transform && cs.transform !== 'none') || (cs.willChange && cs.willChange.includes('transform')) || (cs.filter && cs.filter !== 'none') || (cs.contain && cs.contain !== 'none');
            const clips = /(auto|scroll|hidden|clip)/.test(cs.overflowX) || /(auto|scroll|hidden|clip)/.test(cs.overflowY);
            if (hasTrap || clips) { n.classList.add('doc-unclip'); a++; }
            n = n.parentElement; h++;
        }
    }

    function computeInsets() {
        const vw = Math.max(320, innerWidth), vh = Math.max(320, innerHeight);
        let top = 0, left = 0, right = 0, bottom = 0;
        const hdr = $('[data-app-header]'); const sdr = $('[data-app-sidebar]'); const ftr = $('footer,[data-app-footer]');
        if (hdr) { const r = hdr.getBoundingClientRect(); if (r.height > 0) top = Math.max(top, Math.round(r.bottom)); }
        if (sdr) { const r = sdr.getBoundingClientRect(); if (r.width > 0 && r.left <= 1) left = Math.max(left, Math.round(r.right)); }
        if (ftr) { const r = ftr.getBoundingClientRect(); if (r.height > 0) bottom = Math.max(bottom, Math.round(vh - r.top)); }
        const sdr2 = document.querySelector('aside,.sidebar,[class*="sidebar"],[class*="SideBar"]');
        if (sdr2) { const r = sdr2.getBoundingClientRect(); if (r.width > 0 && r.left <= 1) left = Math.max(left, Math.round(r.right)); }
        top = Math.min(top, Math.floor(vh * 0.5)); left = Math.min(left, Math.floor(vw * 0.6)); right = Math.min(right, Math.floor(vw * 0.4)); bottom = Math.min(bottom, Math.floor(vh * 0.5));
        return { top, left, right, bottom };
    }

    function placeContainer() {
        const host = document.getElementById('doc-scroll');
        if (!host) return;
        const { top, left, right, bottom } = computeInsets();
        host.style.top = top + 'px'; host.style.left = left + 'px'; host.style.right = right + 'px'; host.style.bottom = bottom + 'px';
        host.style.setProperty('overflow-x', 'auto', 'important');
        host.style.setProperty('overflow-y', 'auto', 'important');
        const vInner = host.offsetWidth - host.clientWidth;
        host.style.paddingInlineEnd = (vInner > 0 ? vInner + 'px' : '0');
    }

    /* ===== 프리뷰/스타일 보조 ===== */
    function readJson(id) {
        try {
            const el = document.getElementById(id);
            if (!el) return {};
            const raw = el.textContent || '';
            const v = JSON.parse(raw);
            return typeof v === 'string' ? JSON.parse(v) : v;
        } catch { return {}; }
    }

    const descriptor = readJson('DescriptorJson');
    const preview = readJson('PreviewJson');
    window.descriptor = descriptor;

    const excelColWidthToPx = W => {
        const w = Number(W);
        if (!isFinite(w) || w < 0) return 0;
        return Math.max(0, Math.floor(w * 7 + 5));
    };
    const ptToPx = pt => ((Number(pt) || 0) * 96 / 72);

    function hasVisibleStyle(st) {
        if (!st) return false;
        if (st.fill && st.fill.bg) return true;
        if (st.font && (st.font.bold || st.font.italic || st.font.underline || st.font.size || st.font.name)) return true;
        const b = st.border || {};
        return !!((b.l && b.l !== 'None') || (b.r && b.r !== 'None') || (b.t && b.t !== 'None') || (b.b && b.b !== 'None'));
    }

    function a1ToRC(a1) {
        if (!a1) return null;
        const m = String(a1).toUpperCase().match(/^([A-Z]+)(\d+)$/);
        if (!m) return null;
        const letters = m[1]; const row = parseInt(m[2], 10);
        let col = 0; for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
        return { r: row, c: col };
    }

    const posToMeta = new Map();
    (Array.isArray(descriptor?.inputs) ? descriptor.inputs : []).forEach(f => {
        if (!f?.key) return;
        const rc = a1ToRC(f.a1 || '');
        if (!rc) return;
        posToMeta.set(`${rc.r},${rc.c}`, { key: String(f.key), type: String(f.type || 'Text'), rc });
    });

    const payloadInputs = {};
    window.payloadInputs = payloadInputs;

    function applyStyleToCell(td, cell, st) {
        if (!st || typeof st !== 'object') return;

        if (st.font) {
            const f = st.font;

            // 서버 JSON 어떤 이름이 와도 다 받도록 통합
            const fSize =
                f.size ??
                f.FontSize ??
                f.fontSize ??
                f.sz ??
                f.height;

            // 폰트 패밀리는 웹 공통 폰트로 통일하고 싶으면 name은 무시
            // if (f.name) cell.style.fontFamily = f.name;  // ← 필요 없으면 주석 그대로 두셔도 됩니다.

            if (fSize) {
                cell.style.fontSize = ptToPx(fSize) + 'px';
            }

            if (f.bold) cell.style.fontWeight = '700';
            if (f.italic) cell.style.fontStyle = 'italic';
            if (f.underline) cell.style.textDecoration = 'underline';
        }
        // Excel 폰트(st.font)는 무시하고, eb.doc.preview.css에서 지정한 웹 폰트만 사용
        // (굵기/기울임/크기 모두 CSS 기준 유지)

        if (st.align) {
            const h = String(st.align.h || '').toLowerCase();
            cell.classList.remove('ta-left', 'ta-center', 'ta-right');
            if (h === 'center') cell.classList.add('ta-center');
            else if (h === 'right') cell.classList.add('ta-right');
            else cell.classList.add('ta-left');

            const v = String(st.align.v || '').toLowerCase();
            td.classList.remove('va-top', 'va-middle', 'va-bottom');
            if (v === 'top') td.classList.add('va-top');
            else if (v === 'center' || v === 'middle') td.classList.add('va-middle');
            else td.classList.add('va-bottom');

            if (st.align.wrap) cell.classList.add('wrap');
        }

        const cssOf = s => {
            s = String(s || '').toLowerCase();
            if (!s || s === 'none') return { w: 0, sty: 'none' };
            if (s.includes('double')) return { w: 3, sty: 'double' };
            if (s.includes('mediumdashdotdot') || s.includes('mediumdashdot') || s.includes('mediumdashed')) return { w: 2, sty: 'dashed' };
            if (s.includes('dashdotdot') || s.includes('dashdot') || s.includes('dashed')) return { w: 1, sty: 'dashed' };
            if (s.includes('dotted') || s.includes('hair')) return { w: 1, sty: 'dotted' };
            if (s.includes('thick')) return { w: 3, sty: 'solid' };
            if (s.includes('medium')) return { w: 2, sty: 'solid' };
            return { w: 1, sty: 'solid' };
        };

        const color = '#000';
        if (st.border) {
            const L = cssOf(st.border.l), R = cssOf(st.border.r), T = cssOf(st.border.t), B = cssOf(st.border.b);
            td.style.borderLeft = L.w ? `${L.w}px ${L.sty} ${color}` : 'none';
            td.style.borderRight = R.w ? `${R.w}px ${R.sty} ${color}` : 'none';
            td.style.borderTop = T.w ? `${T.w}px ${T.sty} ${color}` : 'none';
            td.style.borderBottom = B.w ? `${B.w}px ${B.sty} ${color}` : 'none';
        }
    }

    function measureEffectiveContentWidth(tbl) {
        if (!tbl) return 0;
        const baseLeft = tbl.getBoundingClientRect().left;
        let maxRight = 0;
        for (const td of tbl.querySelectorAll('td')) {
            const cs = getComputedStyle(td);
            const hasBorder = ((parseFloat(cs.borderLeftWidth) || 0) > 0) || ((parseFloat(cs.borderRightWidth) || 0) > 0) || ((parseFloat(cs.borderTopWidth) || 0) > 0) || ((parseFloat(cs.borderBottomWidth) || 0) > 0);
            const hasText = (td.textContent || '').trim().length > 0;
            const isEditable = td.classList.contains('eb-editable');
            if (hasBorder || hasText || isEditable) {
                const r = td.getBoundingClientRect().right;
                if (r > maxRight) maxRight = r;
            }
        }
        const eff = Math.ceil(maxRight - baseLeft);
        return (isFinite(eff) && eff > 0) ? eff : 0;
    }

    function computeFinalDocW() {
        const tbl = document.querySelector('#xhost table');
        if (!tbl) return 0;
        const eff = Math.max(1, Number(window.__DOC_EFFW__) || 0, measureEffectiveContentWidth(tbl) | 0);
        const tblW = Math.max(1, Number(window.__DOC_TOTALW__) || 0, tbl.scrollWidth | 0, tbl.offsetWidth | 0);
        return Math.min(eff, tblW);
    }

    function applyFinalWidth() {
        const w = computeFinalDocW();
        const xhost = document.getElementById('xhost');
        if (xhost) xhost.style.width = (w > 0 ? w.toFixed(2) : '1') + 'px';
        window.__DOC_FINALW__ = w;
    }

    let HMAX = 0;
    function updateClampBounds() {
        const sc = document.getElementById('doc-scroll');
        const finalW = computeFinalDocW();
        if (!sc || !finalW) return;
        const viewW = sc.clientWidth | 0;
        const EPS = 6;
        HMAX = Math.max(0, (finalW - viewW - EPS) | 0);
        if (sc.scrollLeft > HMAX) sc.scrollLeft = HMAX;
        if (sc.scrollLeft < 0) sc.scrollLeft = 0;
        window.__DOC_FINALW__ = finalW;
    }

    /* ===== 포커스 링 ===== */
    const cellDivByKey = k => document.querySelector(`#xhost td[data-key="${CSS.escape(k)}"] .cellc`);
    const tdByKey = k => document.querySelector(`#xhost td[data-key="${CSS.escape(k)}"]`);

    function getRingHost() { return document.getElementById('xhost'); }
    function getOrCreateRing() {
        const host = getRingHost(); if (!host) return null;
        let el = host.querySelector(':scope > .eb-group-ring');
        if (!el) { el = document.createElement('div'); el.className = 'eb-group-ring'; host.appendChild(el); }
        return el;
    }
    function showRingForRect(rect) {
        const ring = getOrCreateRing(); if (!ring || !rect) return;
        const pad = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--eb-ring-pad')) || 2;
        ring.style.left = Math.round(rect.left - pad) + 'px';
        ring.style.top = Math.round(rect.top - pad) + 'px';
        ring.style.width = Math.max(0, Math.round(rect.width + pad * 2)) + 'px';
        ring.style.height = Math.max(0, Math.round(rect.height + pad * 2)) + 'px';
        ring.style.display = 'block';
    }
    function hideRingIfIdle() {
        const host = getRingHost(); if (!host) return;
        const hasActiveBlock = host.querySelector('.eb-block[data-active="1"]');
        const hasActiveInput = host.querySelector('td.eb-group input:focus, td.eb-group textarea:focus');
        if (!hasActiveBlock && !hasActiveInput) {
            const ring = getOrCreateRing(); if (ring) ring.style.display = 'none';
        }
    }

    /* ===== 텍스트 폭 클램프 ===== */
    const __eb_canvas = document.createElement('canvas');
    const __eb_ctx = __eb_canvas.getContext('2d');

    function __eb_getAvailWidth(ta) {
        const cs = getComputedStyle(ta);
        const pad = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0);
        const brd = (parseFloat(cs.borderLeftWidth) || 0) + (parseFloat(cs.borderRightWidth) || 0);
        return ta.clientWidth - pad - brd;
    }
    function __eb_ctxFontFrom(ta) {
        const cs = getComputedStyle(ta);
        // letter-spacing 은 canvas 측정에 직접 반영 불가 → CSS에서 normal 고정
        return `${cs.fontStyle} ${cs.fontVariant} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    }

    function __eb_normalizeClamp(ta, maxLines) {
        const norm = s => String(s || '').replace(/\r\n?/g, '\n');
        const original = norm(ta.value);
        const availW = __eb_getAvailWidth(ta);
        __eb_ctx.font = __eb_ctxFontFrom(ta);

        const tokens = [];
        for (let i = 0; i < original.length;) {
            const ch = original[i];
            if (ch === '\n') { tokens.push('\n'); i++; continue; }
            if (ch === ' ') { let j = i + 1; while (j < original.length && original[j] === ' ') j++; tokens.push(original.slice(i, j)); i = j; continue; }
            let j = i + 1; while (j < original.length && original[j] !== '\n' && original[j] !== ' ') j++; tokens.push(original.slice(i, j)); i = j;
        }

        const lines = ['']; let li = 0;
        function put(tok) {
            if (tok === '\n') { if (lines.length < maxLines) { lines.push(''); li++; return true; } return false; }
            const cand = (lines[li] || '') + tok;
            if (__eb_ctx.measureText(cand).width <= availW) { lines[li] = cand; return true; }
            if (lines.length >= maxLines) return false;
            lines.push(''); li++;
            if (__eb_ctx.measureText(tok).width <= availW) { lines[li] = tok; return true; }
            let acc = '';
            for (const c of tok) {
                const t = acc + c;
                if (__eb_ctx.measureText(t).width > availW) {
                    if (lines.length >= maxLines) break;
                    lines[li] = acc; lines.push(''); li++; acc = __eb_ctx.measureText(c).width > availW ? '' : c;
                } else { acc = t; }
            }
            if (acc) lines[li] = acc;
            return lines.length <= maxLines;
        }
        for (const t of tokens) { if (!put(t)) break; }

        const out = (lines.length >= maxLines ? lines.slice(0, maxLines) : lines.concat(Array(maxLines - lines.length).fill(''))).join('\n');
        const pos = Math.min(out.length, (ta.selectionStart ?? out.length));
        ta.value = out; ta.setSelectionRange(pos, pos);
    }

    function __eb_bindBeforeInputWidthClamp(ta, maxLines) {
        const norm = v => String(v || '').replace(/\r\n?/g, '\n');
        const blink = el => { el.classList.add('eb-limit-blink'); setTimeout(() => el.classList.remove('eb-limit-blink'), 260); };

        function tailHasFree(fromLineExclusive) {
            const arr = norm(ta.value).split('\n');
            for (let i = fromLineExclusive + 1; i < maxLines; i++) {
                const s = (arr[i] ?? '');
                if (s.trim().length === 0) return true;
            }
            return false;
        }
        function caretMeta() {
            const raw = ta.value; const s = ta.selectionStart ?? 0; const e = ta.selectionEnd ?? s;
            const up = norm(raw.slice(0, s)); const curLine = (up.match(/\n/g) || []).length;
            const colInLine = up.length - (up.lastIndexOf('\n') + 1);
            return { raw, s, e, curLine, colInLine };
        }
        const canEnterAt = lineIdx => (lineIdx < maxLines - 1) && tailHasFree(lineIdx);

        if (ta.dataset.ebBind === '1') return;
        ta.dataset.ebBind = '1';

        ta.addEventListener('beforeinput', ev => {
            const type = ev.inputType || '';
            if (type.startsWith('delete')) return;
            if (type === 'insertFromPaste') return;

            const isEnter = (type === 'insertParagraph' || type === 'insertLineBreak');
            const data = isEnter ? '\n' : (ev.data || '');
            if (!data && !/insert|composition/i.test(type)) return;

            const { raw, s, e, curLine, colInLine } = caretMeta();

            if (isEnter) {
                ev.preventDefault();
                if (!canEnterAt(curLine)) { blink(ta); return; }
                ta.value = raw.slice(0, s) + '\n' + raw.slice(e);
                const pos = s + 1; ta.setSelectionRange(pos, pos);
                queueMicrotask(() => __eb_normalizeClamp(ta, maxLines));
                return;
            }

            const lines = norm(raw).split('\n');
            const curLineText = lines[curLine] ?? '';
            const canvas = document.createElement('canvas'); const ctx = canvas.getContext('2d'); const cs = getComputedStyle(ta);
            ctx.font = `${cs.fontStyle} ${cs.fontVariant} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;

            const candidate = curLineText.slice(0, colInLine) + (ev.data || '') + curLineText.slice(colInLine + (e - s));
            const tooWide = ctx.measureText(candidate).width > __eb_getAvailWidth(ta);

            if (tooWide) {
                if (!canEnterAt(curLine)) { ev.preventDefault(); blink(ta); return; }
                ev.preventDefault();
                ta.value = raw.slice(0, s) + '\n' + (ev.data || '') + raw.slice(e);
                const pos = s + 1 + (ev.data || '').length; ta.setSelectionRange(pos, pos);
                queueMicrotask(() => __eb_normalizeClamp(ta, maxLines));
                return;
            }

            const simulated = norm(raw.slice(0, s) + (ev.data || '') + raw.slice(e));
            const lineCount = (simulated.match(/\n/g) || []).length + 1;
            if (lineCount > maxLines) { ev.preventDefault(); blink(ta); }
        }, true);

        ta.addEventListener('keydown', e => {
            if (e.key !== 'Enter') return;
            const raw = ta.value; const s = ta.selectionStart ?? 0;
            const up = raw.slice(0, s).replace(/\r\n?/g, '\n');
            const curLine = (up.match(/\n/g) || []).length;
            if (!canEnterAt(curLine)) { e.preventDefault(); blink(ta); }
        }, true);
    }

    const TAB_SIZE = 4, TAB_SPACES = ' '.repeat(TAB_SIZE);

    function __eb_insertSmart_fromStart(ta, text, maxLines) {
        const norm = s => String(s || '').replace(/\r\n?/g, '\n');
        const raw = norm(ta.value);
        const s = ta.selectionStart ?? 0;
        const startLine = (norm(raw.slice(0, s)).match(/\n/g) || []).length;
        const head = raw.split('\n').slice(0, startLine);
        ta.value = head.join('\n'); if (ta.value && !ta.value.endsWith('\n')) ta.value += '\n';
        ta.setSelectionRange(ta.value.length, ta.value.length);

        text = norm(text).replace(/\t/g, TAB_SPACES);

        const availW = __eb_getAvailWidth(ta);
        __eb_ctx.font = __eb_ctxFontFrom(ta);

        const tokens = [];
        for (let i = 0; i < text.length;) {
            const ch = text[i];
            if (ch === '\n') { tokens.push('\n'); i++; continue; }
            if (ch === ' ') { let j = i + 1; while (j < text.length && text[j] === ' ') j++; tokens.push(text.slice(i, j)); i = j; continue; }
            let j = i + 1; while (j < text.length && text[j] !== '\n' && text[j] !== ' ') j++; tokens.push(text.slice(i, j)); i = j;
        }

        const lines = ['']; let li = 0;
        function push(tok) {
            if (tok === '\n') { if (lines.length >= maxLines) return false; lines.push(''); li++; return true; }
            const cand = (lines[li] || '') + tok;
            if (__eb_ctx.measureText(cand).width <= availW) { lines[li] = cand; return true; }
            if (lines.length >= maxLines) return false;
            lines.push(''); li++;
            if (__eb_ctx.measureText(tok).width <= availW) { lines[li] = tok; return true; }
            for (const c of tok) {
                const t = (lines[li] || '') + c;
                if (__eb_ctx.measureText(t).width > availW) {
                    if (lines.length >= maxLines) return false;
                    lines.push(''); li++;
                    if (__eb_ctx.measureText(c).width > availW) return false;
                    lines[li] = c;
                } else { lines[li] = t; }
            }
            return true;
        }
        for (const t of tokens) { if (!push(t)) break; }

        const body = lines.slice(0, maxLines).join('\n');
        const glue = (ta.value && body) ? '\n' : '';
        ta.value = ta.value + glue + body;
        __eb_normalizeClamp(ta, maxLines);
    }

    function bindTabAsSpaces(ta) {
        ta.addEventListener('keydown', (e) => {
            if (e.key !== 'Tab') return;
            e.preventDefault();
            const s = ta.selectionStart || 0;
            const selEnd = ta.selectionEnd || s;
            const before = ta.value.slice(0, s);
            const after = ta.value.slice(selEnd);
            ta.value = before + ' '.repeat(4) + after;
            const pos = s + 4;
            ta.setSelectionRange(pos, pos);
            ta.dispatchEvent(new Event('input'));
        });
    }

    function yToLineIdx(relY, lineH, maxLines) {
        const h = Math.max(1, lineH | 0);
        const idx = Math.floor(Math.max(0, Math.min(relY, maxLines * h - 1)) / h);
        return Math.max(0, Math.min(maxLines - 1, idx));
    }
    function moveCaretToLineEnd(ta, lineIdx) {
        const lines = ta.value.replace(/\r\n?/g, '\n').split('\n');
        const i = Math.max(0, Math.min(lineIdx, Math.max(0, lines.length - 1)));
        let pos = 0; for (let k = 0; k < i; k++) pos += (lines[k] || '').length + 1;
        pos += (lines[i] || '').length;
        ta.setSelectionRange(pos, pos);
    }

    /* ===== 오버레이 그룹 ===== */
    let blocks = [];

    function copyStyleFromFirstCell(el, fc) {
        const td = fc.closest('td');
        const cs = getComputedStyle(fc);
        const rowPx = parseFloat(td?.dataset.rowpx || '') || parseFloat(cs.lineHeight) || 16;
        el.style.lineHeight = rowPx + 'px';
        el.style.fontFamily = cs.fontFamily;
        el.style.fontSize = cs.fontSize;
        el.style.fontWeight = cs.fontWeight;
        el.style.fontStyle = cs.fontStyle;
        el.style.letterSpacing = 'normal';
        el.style.wordSpacing = 'normal';
        el.style.textAlign = cs.textAlign;
        el.style.paddingTop = '0px';
        el.style.paddingBottom = '0px';
        el.style.paddingLeft = cs.paddingLeft;
        el.style.paddingRight = cs.paddingRight;
    }

    function blockRectOfKeys(keys) {
        const tds = keys.map(tdByKey).filter(Boolean);
        if (!tds.length) return null;
        const host = document.getElementById('xhost').getBoundingClientRect();
        let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
        for (const td of tds) {
            const rc = td.getBoundingClientRect();
            l = Math.min(l, rc.left); t = Math.min(t, rc.top); r = Math.max(r, rc.right); b = Math.max(b, rc.bottom + 1);
        }
        const left = Math.floor(l - host.left);
        const top = Math.floor(t - host.top);
        const right = Math.ceil(r - host.left);
        const bottom = Math.ceil(b - host.top);
        return { left, top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
    }

    function clampToTable(rect) {
        const finalW = Number(window.__DOC_FINALW__) || computeFinalDocW();
        const left = Math.max(0, Math.min(rect.left, finalW));
        const right = Math.max(left, Math.min(rect.left + rect.width, finalW));
        return { left, top: rect.top, width: (right - left), height: rect.height };
    }

    function createGroupsWithEditorOverlay(isWrite) {
        if (!isWrite) {
            blocks.forEach(b => { b.block?.remove(); b.ta?.remove(); });
            blocks = [];
            window.__eb_blocks__ = blocks;
            return;
        }

        const groups = (function (inputs) {
            const by = new Map();
            (inputs || []).forEach(f => {
                if (!f?.key || !f?.a1) return;
                const type = String(f.type || 'Text').toLowerCase();
                if (type !== 'text') return;
                const m = String(f.key).match(/^(.*)_(\d+)$/);
                const rc = a1ToRC(f.a1);
                if (!rc) return;
                const base = m ? m[1] : f.key;
                const sig = `${base}:${rc.c}`;
                if (!by.has(sig)) by.set(sig, []);
                by.get(sig).push({ key: f.key, rc });
            });

            const gs = [];
            by.forEach(list => {
                list.sort((a, b) => a.rc.r - b.rc.r);
                let run = [];
                for (let i = 0; i < list.length; i++) {
                    if (i === 0) { run = [list[i]]; continue; }
                    const prev = list[i - 1].rc.r; const cur = list[i].rc.r;
                    if (cur === prev + 1) run.push(list[i]);
                    else { if (run.length) gs.push({ keys: run.map(x => x.key) }); run = [list[i]]; }
                }
                if (run.length) gs.push({ keys: run.map(x => x.key) });
            });
            return gs;
        })(Array.isArray(descriptor?.inputs) ? descriptor.inputs : []);

        blocks.forEach(b => { b.block?.remove(); b.ta?.remove(); });
        blocks = [];

        const host = document.getElementById('xhost');

        for (const g of groups) {
            const keys = g.keys.filter(k => !!tdByKey(k));
            if (!keys.length) continue;

            for (const k of keys) {
                const td = tdByKey(k);
                if (td) td.classList.add('eb-group');
            }

            const firstCell = cellDivByKey(keys[0]);
            if (!firstCell) continue;

            const initLines = keys.map(k => (cellDivByKey(k)?.textContent || ''));
            keys.forEach(k => { const c = cellDivByKey(k); if (c) c.textContent = ''; });

            const block = document.createElement('div');
            block.className = 'eb-block'; block.setAttribute('tabindex', '-1');

            const ta = document.createElement('textarea');
            ta.className = 'eb-ta'; ta.setAttribute('rows', String(keys.length)); ta.setAttribute('wrap', 'off');

            copyStyleFromFirstCell(block, firstCell);
            copyStyleFromFirstCell(ta, firstCell);

            const placeRect = () => {
                const r = blockRectOfKeys(keys) || { left: 0, top: 0, width: 0, height: 0 };
                const rect = clampToTable({ left: Math.floor(r.left), top: Math.floor(r.top), width: Math.ceil(r.width), height: Math.ceil(r.height + 2) });
                block.style.left = rect.left + 'px'; block.style.top = rect.top + 'px'; block.style.width = rect.width + 'px'; block.style.height = rect.height + 'px';
                ta.style.left = rect.left + 'px'; ta.style.top = rect.top + 'px'; ta.style.width = rect.width + 'px'; ta.style.height = rect.height + 'px';
            };
            placeRect();

            ta.value = initLines.some(t => t && t.length) ? initLines.join('\n') : '';

            const lineH = parseFloat(getComputedStyle(ta).lineHeight) || 16;
            __eb_bindBeforeInputWidthClamp(ta, keys.length);

            function syncPayload() {
                let vis = ta.value.replace(/\r\n?/g, '\n').split('\n');
                if (vis.length < keys.length) vis = vis.concat(Array(keys.length - vis.length).fill(''));
                vis = vis.slice(0, keys.length);
                for (let i = 0; i < keys.length; i++) payloadInputs[keys[i]] = vis[i] ?? '';
            }
            syncPayload();

            ta.addEventListener('paste', (e) => {
                e.preventDefault();
                const text = (e.clipboardData?.getData('text/plain') || '').replace(/\u200B/g, '');
                __eb_insertSmart_fromStart(ta, text, keys.length);
                __eb_normalizeClamp(ta, keys.length);
                syncPayload();
            });
            bindTabAsSpaces(ta);
            ta.addEventListener('input', () => { __eb_normalizeClamp(ta, keys.length); syncPayload(); });
            ta.addEventListener('compositionend', () => { ta.dispatchEvent(new Event('input')); });

            for (const k of keys) {
                const td = tdByKey(k); if (!td) continue;
                td.addEventListener('mousedown', e => {
                    e.preventDefault();
                    ta.focus({ preventScroll: true });
                    const rect = ta.getBoundingClientRect();
                    const relY = Math.max(0, Math.min(rect.height - 1, e.clientY - rect.top));
                    const li = yToLineIdx(relY, lineH, keys.length);
                    moveCaretToLineEnd(ta, li);
                });
            }

            ta.addEventListener('focus', () => {
                const r = blockRectOfKeys(keys); if (!r) return;
                block.dataset.active = '1';
                const rect = clampToTable({ left: Math.floor(r.left), top: Math.floor(r.top), width: Math.ceil(r.width), height: Math.ceil(r.height) });
                showRingForRect(rect);
            }, true);
            ta.addEventListener('blur', () => { block.dataset.active = '0'; hideRingIfIdle(); }, true);

            host.appendChild(block);
            host.appendChild(ta);

            blocks.push({
                block, ta, keys, placeRect, refreshRing: () => {
                    if (block.dataset.active === '1') {
                        const r = blockRectOfKeys(keys); if (!r) return;
                        const rect = clampToTable({ left: Math.floor(r.left), top: Math.floor(r.top), width: Math.ceil(r.width), height: Math.ceil(r.height) });
                        showRingForRect(rect);
                    }
                }
            });
        }

        window.__eb_blocks__ = blocks;
        window.dispatchEvent(new CustomEvent('eb-editor-mounted'));
        addEventListener('resize', () => { blocks.forEach(b => { b.placeRect(); b.refreshRing(); }); });
    }

    /* ===== 프리뷰 마운트 ===== */
    function toIsoDateOrEmpty(text) {
        if (!text) return '';
        const s = String(text).trim();
        let m = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
        if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
        return '';
    }

    function mount(_, p, options) {
        const opts = options || {};
        const isWrite = (typeof opts.isWrite === 'boolean') ? opts.isWrite : __EB_IS_WRITE__;
        const xhost = document.getElementById('xhost');
        if (!xhost || !p || !Array.isArray(p.cells) || !p.cells.length) {
            if (xhost) xhost.innerHTML = '<div class="alert alert-danger">DOC_Err_PreviewFailed</div>';
            return;
        }

        // CSS 변수 미로드/순서 문제 대비: 런타임에서도 폰트 기준을 강제
        //const xlPrev = document.getElementById('xlPreview');
        //if (xlPrev) {
        //    xlPrev.style.fontFamily = `var(--eb-font)`;
        //    xlPrev.style.fontSize = `var(--eb-font-size)`;
        //    xlPrev.style.lineHeight = `var(--eb-line-height)`;
        //    xlPrev.style.letterSpacing = 'normal';
        //    xlPrev.style.wordSpacing = 'normal';
        //}
        const xlPrev = document.getElementById('xlPreview');
        if (xlPrev) {
            xlPrev.style.fontFamily = `var(--eb-font)`;
            xlPrev.style.letterSpacing = 'normal';
            xlPrev.style.wordSpacing = 'normal';
        }

        const styles = p.styles || {};
        const allRows = p.cells.length;
        let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
        const mark = (r, c) => { minR = Math.min(minR, r); maxR = Math.max(maxR, r); minC = Math.min(minC, c); maxC = Math.max(maxC, c); };

        for (let r = 1; r <= allRows; r++) {
            const row = p.cells[r - 1] || [];
            for (let c = 1; c <= row.length; c++) {
                const v = row[c - 1];
                if (v !== '' && v != null) mark(r, c);
                const st = styles[`${r},${c}`];
                if (hasVisibleStyle(st)) mark(r, c);
            }
        }
        (p.merges || []).forEach(m => { const [r1, c1, r2, c2] = m.map(n => parseInt(n, 10) || 0); mark(r1, c1); mark(r2, c2); });
        (Array.isArray(descriptor?.inputs) ? descriptor.inputs : []).forEach(f => { const rc = a1ToRC(f?.a1); if (rc) mark(rc.r, rc.c); });

        if (!isFinite(minR) || !isFinite(minC)) { minR = 1; maxR = 1; minC = 1; maxC = 1; }

        const maxColsFromCells = Math.max(...p.cells.map(r => r.length), 1);
        maxC = Math.min(maxC, maxColsFromCells); minC = Math.max(minC, 1);

        const mergeMap = new Map();
        (p.merges || []).forEach(m => {
            let [r1, c1, r2, c2] = m.map(n => parseInt(n, 10));
            r1 = Math.max(r1, minR); c1 = Math.max(c1, minC);
            r2 = Math.min(r2, maxR); c2 = Math.min(c2, maxC);
            if (r1 > r2 || c1 > c2) return;
            const master = `${r1}-${c1}`;
            mergeMap.set(master, { master: true, rs: r2 - r1 + 1, cs: c2 - c1 + 1 });
            for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
                const k = `${r}-${c}`; if (k !== master) mergeMap.set(k, { covered: true });
            }
        });

        const styleGrid = Array.from({ length: maxR + 1 }, () => Array(maxC + 1).fill(null));
        for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
                const key = `${r},${c}`;
                const st = styles[key] || {};
                const border = Object.assign({ l: 'None', r: 'None', t: 'None', b: 'None' }, st.border || {});
                styleGrid[r][c] = { font: st.font || null, align: st.align || null, fill: st.fill || null, border };
            }
        }

        const weight = s => {
            s = String(s || '').toLowerCase();
            if (!s || s === 'none') return 0;
            if (s.includes('double')) return 6;
            if (s.includes('thick')) return 5;
            if (s.includes('mediumdashdotdot') || s.includes('mediumdashdot') || s.includes('mediumdashed') || s.includes('medium')) return 4;
            if (s.includes('dashed') || s.includes('dashdot') || s.includes('dashdotdot')) return 3;
            if (s.includes('dotted') || s.includes('hair')) return 2;
            return 1;
        };
        const stronger = (a, b) => (weight(a) >= weight(b) ? a : b);
        for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
                const cur = styleGrid[r][c]; if (!cur) continue;
                if (c < maxC) { const right = styleGrid[r][c + 1]; if (right) { const pick = stronger(cur.border.r, right.border.l); cur.border.r = pick; right.border.l = pick; } }
                if (r < maxR) { const down = styleGrid[r + 1][c]; if (down) { const pick = stronger(cur.border.b, down.border.t); cur.border.b = pick; down.border.t = pick; } }
            }
        }

        const colPxAt = c => { const wChar = (p.colW || [])[c - 1]; return excelColWidthToPx(wChar ?? 8.43); };
        const sumColPx = (c1, c2) => { let s = 0; for (let i = c1; i <= c2; i++) s += colPxAt(i); return s; };

        const tbl = document.createElement('table'); tbl.className = 'xlfb';
        const colgroup = document.createElement('colgroup');
        for (let c = minC; c <= maxC; c++) { const cg = document.createElement('col'); cg.style.width = colPxAt(c).toFixed(2) + 'px'; colgroup.appendChild(cg); }
        tbl.appendChild(colgroup);

        const tbody = document.createElement('tbody');
        const rowHeights = Array.isArray(p.rowH) ? p.rowH : [];
        const DEFAULT_ROW_PT = 15;

        for (let r = minR; r <= maxR; r++) {
            const tr = document.createElement('tr');
            const pt = (rowHeights[r - 1] != null) ? rowHeights[r - 1] : DEFAULT_ROW_PT;
            const rowPx = ptToPx(pt);
            tr.style.height = rowPx + 'px';

            for (let c = minC; c <= maxC; c++) {
                const key = `${r}-${c}`;
                const mm = mergeMap.get(key);
                if (mm && mm.covered) continue;

                const td = document.createElement('td');
                td.dataset.rowpx = String(rowPx);

                if (mm && mm.master) { if (mm.rs > 1) td.setAttribute('rowspan', String(mm.rs)); if (mm.cs > 1) td.setAttribute('colspan', String(mm.cs)); }
                if (!(mm && mm.master && mm.cs > 1)) td.style.width = colPxAt(c) + 'px';

                const cell = document.createElement('div');
                cell.className = 'cellc';
                // 셀 표시도 textarea와 동일 라인 기준을 강제
                cell.style.lineHeight = rowPx + 'px';
                if (!mm) cell.style.maxHeight = rowPx + 'px';

                const v = (preview.cells[r - 1]?.[c - 1] ?? '');
                cell.appendChild(document.createTextNode(v === '' ? '' : String(v)));

                applyStyleToCell(td, cell, styleGrid[r][c]);

                const m = posToMeta.get(`${r},${c}`);
                const fieldKey = m?.key;
                const fieldType = (m?.type || 'Text').toLowerCase();
                const editable = !!fieldKey && !(mm && !mm.master) && isWrite;

                if (editable) {
                    td.setAttribute('data-key', fieldKey);
                    td.classList.add('eb-editable');

                    if (fieldType === 'date') {
                        td.dataset.type = 'date';
                        td.classList.add('eb-group');

                        const input = document.createElement('input');
                        input.type = 'date';
                        input.className = 'eb-input-date';

                        // 2025.11.14 Changed: 항상 폼 여는 시점의 "오늘 날짜(로컬 기준)"로 기본 값 설정
                        const today = new Date();
                        const yyyy = today.getFullYear();
                        const mm = String(today.getMonth() + 1).padStart(2, '0');
                        const dd = String(today.getDate()).padStart(2, '0');
                        input.value = `${yyyy}-${mm}-${dd}`;

                        // 셀의 기존 텍스트(엑셀 값)는 무시
                        cell.textContent = '';
                        cell.appendChild(input);

                        const sync = () => { payloadInputs[fieldKey] = input.value || ''; };
                        input.addEventListener('change', sync);
                        input.addEventListener('blur', sync);
                        if (!(fieldKey in payloadInputs)) sync();

                        td.addEventListener('mousedown', (e) => {
                            if (e.button !== 0) return;
                            e.preventDefault();
                            input.focus({ preventScroll: true });
                            if (typeof input.showPicker === 'function') {
                                try { input.showPicker(); } catch { /* no-op */ }
                            }
                        });
                    }
                }

                td.appendChild(cell);
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
        tbl.appendChild(tbody);

        const totalW = sumColPx(minC, maxC);
        tbl.style.width = totalW + 'px';

        const xhostEl = document.getElementById('xhost');
        xhostEl.innerHTML = '';
        xhostEl.appendChild(tbl);

        window.__DOC_TOTALW__ = Math.max(1, totalW | 0);
        window.__DOC_EFFW__ = Math.max(1, measureEffectiveContentWidth(tbl) | 0);

        applyFinalWidth();
        requestAnimationFrame(() => requestAnimationFrame(updateClampBounds));

        // Detail 모드에서는 그룹 표시/편집 오버레이 생성 안 함
        try { if (typeof colorizeInputGroups === 'function') colorizeInputGroups(isWrite); } catch (e) { console.warn('colorizeInputGroups skipped', e); }
        createGroupsWithEditorOverlay(isWrite);
    }

    /* ===== 초기화/관찰 ===== */
    const host = document.getElementById('doc-scroll');
    if (host) {
        escapeFixedTraps(host);
        placeContainer();
        host.addEventListener('scroll', () => {
            if (host.scrollLeft > HMAX) host.scrollLeft = HMAX;
            if (host.scrollLeft < 0) host.scrollLeft = 0;
        }, { passive: true });
    }

    function rerenderAll() {
        placeContainer();
        applyFinalWidth();
        updateClampBounds();
        if (window.__eb_blocks__) {
            window.__eb_blocks__.forEach(b => { b.placeRect?.(); b.refreshRing?.(); });
        }
    }

    async function firstReflow() {
        try { await (document.fonts?.ready || Promise.resolve()); } catch { }
        await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
        rerenderAll();
        setTimeout(rerenderAll, 0);
    }

    try {
        if (preview && preview.cells) {
            mount('#xlPreview', preview, { isWrite: __EB_IS_WRITE__ });
        }
    } catch (e) {
        console.error(e);
        err('DOC_Err_PreviewFailed');
    }

    addEventListener('load', firstReflow, { once: true });
    firstReflow();
    addEventListener('resize', rerenderAll);

    const xhostObsTarget = document.getElementById('xhost');
    if (xhostObsTarget) {
        let moQueued = false;
        new MutationObserver(() => {
            if (moQueued) return;
            moQueued = true;
            requestAnimationFrame(() => {
                const tbl = document.querySelector('#xhost table');
                if (tbl) {
                    window.__DOC_EFFW__ = Math.max(Number(window.__DOC_EFFW__) || 0, measureEffectiveContentWidth(tbl) | 0);
                }
                rerenderAll();
                moQueued = false;
            });
        }).observe(xhostObsTarget, { attributes: true, childList: true, subtree: true });
    }

    /* ===== 전역 노출 ===== */
    window.EBDocPreview = {
        T, info, ok, err,
        descriptor, preview,
        get payloadInputs() { return payloadInputs; },
        mount, rerenderAll, firstReflow, escapeFixedTraps, placeContainer, readJson,
        setWriteMode: function (flag) { __EB_IS_WRITE__ = !!flag; window.__DOC_IS_WRITE__ = __EB_IS_WRITE__; },
        get isWrite() { return __EB_IS_WRITE__; },
        getMetrics: () => ({ totalW: window.__DOC_TOTALW__ || 0, effW: window.__DOC_EFFW__ || 0, finalW: window.__DOC_FINALW__ || 0, hMax: HMAX })
    };
})();