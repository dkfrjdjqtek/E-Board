// 2025.11.20 Changed: Compose 쓰기 모드에서 xhost 폭을 테이블 scrollWidth 기준으로 고정해 가로 스크롤 끝에서 잘리지 않도록 수정 및 가로 클램프 롤백 완전 제거
// 2025.11.21 Changed: EBPreview 구 스키마(border/align) 호환 추가 및 '없는 보더'는 건드리지 않도록 수정, 폰트 크기 적용 제거(공통 CSS 기준 유지)
// 2025.11.14 Changed: Excel 폰트 무시하고 eb.doc.preview.css 기준 폰트 통일(applyStyleToCell 폰트 적용 제거)
// 2025.11.14 Changed: colorizeInputGroups 폴백 추가 및 mount 안전화 프리뷰 실패 방지
// 2025.11.14 Changed: Compose Detail 폰트 px 고정과 동일 줄바꿈 일치 위해 측정폭 계산에 border 포함 cellc lineHeight=rowPx 강제 copyStyle 자간 normal 고정 hideRingIfIdle 안정화
// 2025.11.14 Added: DOC Compose Detail 공통 프리뷰 렌더링 스크립트 분리
// 2025.11.26 Changed DocTLMap 미리보기 구동 오류 수정 const 사용 구문 에러 제거 및 ES5 호환 재작성

(function () {
    // colorizeInputGroups 폴백
    if (typeof window.colorizeInputGroups !== "function") {
        window.colorizeInputGroups = function () { };
    }

    // 쓰기 모드 플래그 (기본 true)
    var __EB_IS_WRITE__ = (typeof window.__DOC_IS_WRITE__ === "boolean") ? !!window.__DOC_IS_WRITE__ : true;
    window.__DOC_IS_WRITE__ = __EB_IS_WRITE__;

    /* ================= 공통 유틸 ================= */

    function $(q, root) {
        return (root || document).querySelector(q);
    }
    function $all(q, root) {
        return Array.prototype.slice.call((root || document).querySelectorAll(q));
    }
    function $alert() {
        return document.getElementById("doc-alert");
    }

    function decodeHtmlEntities(str) {
        if (str == null) return "";
        var s = String(str);
        s = s.replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCodePoint(parseInt(h, 16)); });
        s = s.replace(/&#(\d+);/g, function (_, d) { return String.fromCodePoint(parseInt(d, 10)); });
        s = s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
        return s;
    }

    function T(m) {
        if (m == null) return "";
        if (Object.prototype.toString.call(m) === "[object Array]") {
            var arr = [];
            for (var i = 0; i < m.length; i++) arr.push(T(m[i]));
            return arr.join("\n");
        }
        var s = String(m);
        if (window.__RESX && Object.prototype.hasOwnProperty.call(window.__RESX, s)) {
            return decodeHtmlEntities(window.__RESX[s]);
        }
        return decodeHtmlEntities(s);
    }

    function info(m) {
        var host = $alert();
        if (!host) return;
        host.innerHTML = "";
        var div = document.createElement("div");
        div.className = "alert alert-info";
        div.textContent = T(m);
        host.appendChild(div);
    }
    function ok(m) {
        var host = $alert();
        if (!host) return;
        host.innerHTML = "";
        var div = document.createElement("div");
        div.className = "alert alert-success";
        div.textContent = T(m);
        host.appendChild(div);
    }
    function err(m) {
        var host = $alert();
        if (!host) return;
        host.innerHTML = "";
        var div = document.createElement("div");
        div.className = "alert alert-danger";
        div.textContent = T(m || "DOC_Err_PreviewFailed");
        host.appendChild(div);
    }

    /* ================= 레이아웃(스크롤 호스트) ================= */

    function escapeFixedTraps(el, maxHops, maxApply) {
        maxHops = maxHops || 10;
        maxApply = maxApply || 4;
        var n = el.parentElement;
        var hops = 0;
        var applied = 0;
        while (n && hops < maxHops && applied < maxApply) {
            var cs = window.getComputedStyle(n);
            var hasTrap =
                (cs.transform && cs.transform !== "none") ||
                (cs.willChange && cs.willChange.indexOf("transform") >= 0) ||
                (cs.filter && cs.filter !== "none") ||
                (cs.contain && cs.contain !== "none");
            var clips =
                /(auto|scroll|hidden|clip)/.test(cs.overflowX) ||
                /(auto|scroll|hidden|clip)/.test(cs.overflowY);

            if (hasTrap || clips) {
                if (!n.classList.contains("doc-unclip")) {
                    n.classList.add("doc-unclip");
                }
                applied++;
            }

            n = n.parentElement;
            hops++;
        }
    }

    function computeInsets() {
        var vw = Math.max(320, window.innerWidth || 0);
        var vh = Math.max(320, window.innerHeight || 0);
        var top = 0, left = 0, right = 0, bottom = 0;

        var hdr = document.querySelector("[data-app-header]");
        var sdr = document.querySelector("[data-app-sidebar]");
        var ftr = document.querySelector("footer,[data-app-footer]");

        if (hdr) {
            var r1 = hdr.getBoundingClientRect();
            if (r1.height > 0) top = Math.max(top, Math.round(r1.bottom));
        }
        if (sdr) {
            var r2 = sdr.getBoundingClientRect();
            if (r2.width > 0 && r2.left <= 1) left = Math.max(left, Math.round(r2.right));
        }
        if (ftr) {
            var r3 = ftr.getBoundingClientRect();
            if (r3.height > 0) bottom = Math.max(bottom, Math.round(vh - r3.top));
        }

        var sdr2 = document.querySelector("aside,.sidebar,[class*='sidebar'],[class*='SideBar']");
        if (sdr2) {
            var r4 = sdr2.getBoundingClientRect();
            if (r4.width > 0 && r4.left <= 1) left = Math.max(left, Math.round(r4.right));
        }

        top = Math.min(top, Math.floor(vh * 0.5));
        left = Math.min(left, Math.floor(vw * 0.6));
        right = Math.min(right, Math.floor(vw * 0.4));
        bottom = Math.min(bottom, Math.floor(vh * 0.5));

        return { top: top, left: left, right: right, bottom: bottom };
    }

    function placeContainer() {
        var host = document.getElementById("doc-scroll");
        if (!host) return;

        var insets = computeInsets();
        host.style.position = host.style.position || "relative";
        host.style.top = insets.top + "px";
        host.style.left = insets.left + "px";
        host.style.right = insets.right + "px";
        host.style.bottom = insets.bottom + "px";
        host.style.overflowX = "auto";
        host.style.overflowY = "auto";
        host.style.paddingInlineEnd = "0px";
    }

    /* ================= JSON 로딩 ================= */

    function readJson(id) {
        try {
            var el = document.getElementById(id);
            if (!el) return {};
            var raw = el.textContent || "";
            var v = JSON.parse(raw);
            if (typeof v === "string") return JSON.parse(v);
            return v || {};
        } catch (e) {
            return {};
        }
    }

    var descriptor = readJson("DescriptorJson");
    var preview = readJson("PreviewJson");
    window.descriptor = descriptor;

    /* ================= 좌표/A1 유틸 ================= */

    function a1ToRC(a1) {
        if (!a1) return null;
        var m = String(a1).toUpperCase().match(/^([A-Z]+)(\d+)$/);
        if (!m) return null;
        var letters = m[1];
        var row = parseInt(m[2], 10);
        var col = 0;
        for (var i = 0; i < letters.length; i++) {
            col = col * 26 + (letters.charCodeAt(i) - 64);
        }
        return { r: row, c: col };
    }

    var posToMeta = new window.Map ? new Map() : null;
    var payloadInputs = {};
    window.payloadInputs = payloadInputs;

    (function buildPosMeta() {
        if (!posToMeta) return;
        var list = (descriptor && descriptor.inputs && descriptor.inputs.length) ? descriptor.inputs : [];
        for (var i = 0; i < list.length; i++) {
            var f = list[i];
            if (!f || !f.key) continue;
            var rc = a1ToRC(f.a1 || "");
            if (!rc) continue;
            posToMeta.set(rc.r + "," + rc.c, {
                key: String(f.key),
                type: String(f.type || "Text"),
                rc: rc
            });
        }
    })();

    /* ================= 치수 변환 ================= */

    function excelColWidthToPx(W) {
        var w = Number(W);
        if (!isFinite(w) || w < 0) return 0;
        return Math.max(0, Math.floor(w * 7 + 5));
    }
    function ptToPx(pt) {
        var n = Number(pt) || 0;
        return n * 96 / 72;
    }

    /* ================= 스타일/보더 적용 ================= */

    function hasVisibleStyle(st) {
        if (!st) return false;
        if (st.fill && st.fill.bg) return true;
        if (st.font && (st.font.bold || st.font.italic || st.font.underline || st.font.size || st.font.name)) return true;
        var b = st.border || {};
        return !!((b.l && b.l !== "None") || (b.r && b.r !== "None") || (b.t && b.t !== "None") || (b.b && b.b !== "None"));
    }

    function applyStyleToCell(td, cell, st) {
        if (!st || typeof st !== "object") return;

        if (st.font) {
            var f = st.font;
            var fSize = f.size || f.FontSize || f.fontSize || f.sz || f.height;
            if (fSize) {
                cell.style.fontSize = ptToPx(fSize) + "px";
            }
            if (f.bold) cell.style.fontWeight = "700";
            if (f.italic) cell.style.fontStyle = "italic";
            if (f.underline) cell.style.textDecoration = "underline";
        }

        if (st.align) {
            var h = String(st.align.h || "").toLowerCase();
            cell.classList.remove("ta-left", "ta-center", "ta-right");
            if (h === "center") cell.classList.add("ta-center");
            else if (h === "right") cell.classList.add("ta-right");
            else cell.classList.add("ta-left");

            var v = String(st.align.v || "").toLowerCase();
            td.classList.remove("va-top", "va-middle", "va-bottom");
            if (v === "top") td.classList.add("va-top");
            else if (v === "center" || v === "middle") td.classList.add("va-middle");
            else td.classList.add("va-bottom");

            if (st.align.wrap) cell.classList.add("wrap");
        }

        function cssOf(s) {
            s = String(s || "").toLowerCase();
            if (!s || s === "none") return { w: 0, sty: "none" };
            if (s.indexOf("double") >= 0) return { w: 3, sty: "double" };
            if (s.indexOf("mediumdashdotdot") >= 0 || s.indexOf("mediumdashdot") >= 0 || s.indexOf("mediumdashed") >= 0) return { w: 2, sty: "dashed" };
            if (s.indexOf("dashdotdot") >= 0 || s.indexOf("dashdot") >= 0 || s.indexOf("dashed") >= 0) return { w: 1, sty: "dashed" };
            if (s.indexOf("dotted") >= 0 || s.indexOf("hair") >= 0) return { w: 1, sty: "dotted" };
            if (s.indexOf("thick") >= 0) return { w: 3, sty: "solid" };
            if (s.indexOf("medium") >= 0) return { w: 2, sty: "solid" };
            return { w: 1, sty: "solid" };
        }

        //var color = "#000";
        //if (st.border) {
        //    var L = cssOf(st.border.l);
        //    var R = cssOf(st.border.r);
        //    var T2 = cssOf(st.border.t);
        //    var B = cssOf(st.border.b);

        //    td.style.borderLeft = L.w ? (L.w + "px " + L.sty + " " + color) : "none";
        //    td.style.borderRight = R.w ? (R.w + "px " + R.sty + " " + color) : "none";
        //    td.style.borderTop = T2.w ? (T2.w + "px " + T2.sty + " " + color) : "none";
        //    td.style.borderBottom = B.w ? (B.w + "px " + B.sty + " " + color) : "none";
        //}
        const color = '#000';
        if (st.border) {
            const L = cssOf(st.border.l);
            const R = cssOf(st.border.r);
            const T = cssOf(st.border.t);
            const B = cssOf(st.border.b);

            // ★ w>0 인 경우에만 인라인 스타일로 덮어쓴다.
            if (L.w) td.style.borderLeft = `${L.w}px ${L.sty} ${color}`;
            if (R.w) td.style.borderRight = `${R.w}px ${R.sty} ${color}`;
            if (T.w) td.style.borderTop = `${T.w}px ${T.sty} ${color}`;
            if (B.w) td.style.borderBottom = `${B.w}px ${B.sty} ${color}`;
        }
    }

    /* ================= 폭 계산 (문서 끝) ================= */

    function measureEffectiveContentWidth(tbl) {
        if (!tbl) return 0;
        var rectBase = tbl.getBoundingClientRect();
        var baseLeft = rectBase.left;
        var maxRight = 0;

        var tds = tbl.querySelectorAll("td");
        for (var i = 0; i < tds.length; i++) {
            var td = tds[i];
            var cs = window.getComputedStyle(td);
            var hasBorder =
                ((parseFloat(cs.borderLeftWidth) || 0) > 0) ||
                ((parseFloat(cs.borderRightWidth) || 0) > 0) ||
                ((parseFloat(cs.borderTopWidth) || 0) > 0) ||
                ((parseFloat(cs.borderBottomWidth) || 0) > 0);
            var hasText = (td.textContent || "").replace(/\s+/g, "").length > 0;
            var isEditable = td.classList.contains("eb-editable");

            if (hasBorder || hasText || isEditable) {
                var r = td.getBoundingClientRect();
                if (r.right > maxRight) maxRight = r.right;
            }
        }

        var eff = Math.ceil(maxRight - baseLeft);
        if (!isFinite(eff) || eff <= 0) return 0;
        return eff;
    }

    var HMAX = 0;
    var H_MARGIN = 12;

    function computeFinalDocW() {
        var tbl = document.querySelector("#xhost table");
        if (!tbl) return 0;

        var eff = Math.max(
            1,
            Number(window.__DOC_EFFW__ || 0) || 0,
            measureEffectiveContentWidth(tbl) | 0
        );

        var margin = 12;
        return eff + margin;
    }

    function applyFinalWidth() {
        var tbl = document.querySelector("#xhost table");
        if (!tbl) return;

        var w = computeFinalDocW();
        var xhost = document.getElementById("xhost");
        if (xhost) {
            xhost.style.width = (w > 0 ? w.toFixed(2) : "1") + "px";
        }
        window.__DOC_FINALW__ = w;
        setTimeout(function () { debugHorizontalBounds("after applyFinalWidth"); }, 0);
    }

    function updateClampBounds() {
        var host = document.getElementById("doc-scroll");
        if (!host) {
            HMAX = 0;
            window.__DOC_HMAX__ = 0;
            return;
        }

        var finalW = Number(window.__DOC_FINALW__ || 0) || 0;
        var clientW = host.clientWidth || 0;

        if (!finalW || finalW <= clientW) {
            HMAX = 0;
            window.__DOC_HMAX__ = 0;
        } else {
            HMAX = finalW - clientW + H_MARGIN;
            if (HMAX < 0) HMAX = 0;
            window.__DOC_HMAX__ = HMAX;
        }

        setTimeout(function () { debugHorizontalBounds("after updateClampBounds"); }, 0);
    }

    /* ================= 포커스 링 / 블록 편집 (Compose용) ================= */

    function getRingHost() {
        return document.getElementById("xhost");
    }
    function getOrCreateRing() {
        var host = getRingHost();
        if (!host) return null;
        var el = host.querySelector(":scope > .eb-group-ring");
        if (!el) {
            el = document.createElement("div");
            el.className = "eb-group-ring";
            host.appendChild(el);
        }
        return el;
    }

    function showRingForRect(rect) {
        var ring = getOrCreateRing();
        if (!ring || !rect) return;

        var pad = parseFloat(window.getComputedStyle(document.documentElement).getPropertyValue("--eb-ring-pad")) || 2;
        ring.style.left = Math.round(rect.left - pad) + "px";
        ring.style.top = Math.round(rect.top - pad) + "px";
        ring.style.width = Math.max(0, Math.round(rect.width + pad * 2)) + "px";
        ring.style.height = Math.max(0, Math.round(rect.height + pad * 2)) + "px";
        ring.style.display = "block";
    }

    function hideRingIfIdle() {
        var host = getRingHost();
        if (!host) return;
        var hasBlock = host.querySelector(".eb-block[data-active='1']");
        var hasInput = host.querySelector("td.eb-group input:focus, td.eb-group textarea:focus");
        if (!hasBlock && !hasInput) {
            var ring = getOrCreateRing();
            if (ring) ring.style.display = "none";
        }
    }

    function cellDivByKey(k) {
        try {
            return document.querySelector('#xhost td[data-key="' + CSS.escape(k) + '"] .cellc');
        } catch (e) {
            // CSS.escape 미지원 브라우저 대비
            return document.querySelector('#xhost td[data-key="' + k + '"] .cellc');
        }
    }
    function tdByKey(k) {
        try {
            return document.querySelector('#xhost td[data-key="' + CSS.escape(k) + '"]');
        } catch (e) {
            return document.querySelector('#xhost td[data-key="' + k + '"]');
        }
    }

    function blockRectOfKeys(keys) {
        var host = document.getElementById("xhost");
        if (!host) return null;

        var tds = [];
        for (var i = 0; i < keys.length; i++) {
            var td = tdByKey(keys[i]);
            if (td) tds.push(td);
        }
        if (!tds.length) return null;

        var hostRect = host.getBoundingClientRect();
        var l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;

        for (var j = 0; j < tds.length; j++) {
            var rc = tds[j].getBoundingClientRect();
            if (rc.left < l) l = rc.left;
            if (rc.top < t) t = rc.top;
            if (rc.right > r) r = rc.right;
            if (rc.bottom > b) b = rc.bottom + 1;
        }

        var left = Math.floor(l - hostRect.left);
        var top = Math.floor(t - hostRect.top);
        var right = Math.ceil(r - hostRect.left);
        var bottom = Math.ceil(b - hostRect.top);

        return {
            left: left,
            top: top,
            width: Math.max(0, right - left),
            height: Math.max(0, bottom - top)
        };
    }

    function clampToTable(rect) {
        var finalW = Number(window.__DOC_FINALW__ || 0) || computeFinalDocW();
        var left = Math.max(0, Math.min(rect.left, finalW));
        var right = Math.max(left, Math.min(rect.left + rect.width, finalW));
        return {
            left: left,
            top: rect.top,
            width: (right - left),
            height: rect.height
        };
    }

    var blocks = [];

    function copyStyleFromFirstCell(el, fc) {
        var td = fc.closest ? fc.closest("td") : fc.parentNode;
        var cs = window.getComputedStyle(fc);
        var rowPx = parseFloat(td && td.getAttribute("data-rowpx")) || parseFloat(cs.lineHeight) || 16;

        el.style.lineHeight = rowPx + "px";
        el.style.fontFamily = cs.fontFamily;
        el.style.fontSize = cs.fontSize;
        el.style.fontWeight = cs.fontWeight;
        el.style.fontStyle = cs.fontStyle;
        el.style.letterSpacing = "normal";
        el.style.wordSpacing = "normal";
        el.style.textAlign = cs.textAlign;
        el.style.paddingTop = "0px";
        el.style.paddingBottom = "0px";
        el.style.paddingLeft = cs.paddingLeft;
        el.style.paddingRight = cs.paddingRight;
    }

    /* ===== textarea 폭/줄수 제한 (Compose) ===== */

    var __eb_canvas = document.createElement("canvas");
    var __eb_ctx = __eb_canvas.getContext("2d");

    function __eb_getAvailWidth(ta) {
        var cs = window.getComputedStyle(ta);
        var pad = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0);
        var brd = (parseFloat(cs.borderLeftWidth) || 0) + (parseFloat(cs.borderRightWidth) || 0);
        return ta.clientWidth - pad - brd;
    }
    function __eb_ctxFontFrom(ta) {
        var cs = window.getComputedStyle(ta);
        return [cs.fontStyle, cs.fontVariant, cs.fontWeight, cs.fontSize, cs.fontFamily].join(" ");
    }

    function __eb_normalizeClamp(ta, maxLines) {
        function norm(s) {
            return String(s || "").replace(/\r\n?/g, "\n");
        }
        var original = norm(ta.value);
        var availW = __eb_getAvailWidth(ta);
        __eb_ctx.font = __eb_ctxFontFrom(ta);

        var tokens = [];
        var i, j, ch;
        for (i = 0; i < original.length;) {
            ch = original[i];
            if (ch === "\n") {
                tokens.push("\n"); i++; continue;
            }
            if (ch === " ") {
                j = i + 1;
                while (j < original.length && original[j] === " ") j++;
                tokens.push(original.slice(i, j));
                i = j; continue;
            }
            j = i + 1;
            while (j < original.length && original[j] !== "\n" && original[j] !== " ") j++;
            tokens.push(original.slice(i, j));
            i = j;
        }

        var lines = [""];
        var li = 0;

        function put(tok) {
            if (tok === "\n") {
                if (lines.length < maxLines) {
                    lines.push("");
                    li++;
                    return true;
                }
                return false;
            }
            var cur = lines[li] || "";
            var cand = cur + tok;
            if (__eb_ctx.measureText(cand).width <= availW) {
                lines[li] = cand;
                return true;
            }
            if (lines.length >= maxLines) return false;

            lines.push("");
            li++;
            if (__eb_ctx.measureText(tok).width <= availW) {
                lines[li] = tok;
                return true;
            }
            var acc = "";
            for (var k = 0; k < tok.length; k++) {
                var c = tok.charAt(k);
                var t = acc + c;
                if (__eb_ctx.measureText(t).width > availW) {
                    if (lines.length >= maxLines) break;
                    lines[li] = acc;
                    lines.push("");
                    li++;
                    acc = __eb_ctx.measureText(c).width > availW ? "" : c;
                } else {
                    acc = t;
                }
            }
            if (acc) lines[li] = acc;
            return lines.length <= maxLines;
        }

        for (i = 0; i < tokens.length; i++) {
            if (!put(tokens[i])) break;
        }

        if (lines.length > maxLines) lines = lines.slice(0, maxLines);
        while (lines.length < maxLines) lines.push("");

        var out = lines.join("\n");
        var pos = Math.min(out.length, (typeof ta.selectionStart === "number" ? ta.selectionStart : out.length));
        ta.value = out;
        ta.setSelectionRange(pos, pos);
    }

    function __eb_bindBeforeInputWidthClamp(ta, maxLines) {
        function norm(s) { return String(s || "").replace(/\r\n?/g, "\n"); }

        function blink(el) {
            if (!el) return;
            el.classList.add("eb-limit-blink");
            setTimeout(function () { el.classList.remove("eb-limit-blink"); }, 260);
        }

        function tailHasFree(fromLineExclusive) {
            var arr = norm(ta.value).split("\n");
            for (var i = fromLineExclusive + 1; i < maxLines; i++) {
                var s = arr[i] || "";
                if (s.replace(/\s+/g, "").length === 0) return true;
            }
            return false;
        }

        function caretMeta() {
            var raw = ta.value;
            var s = typeof ta.selectionStart === "number" ? ta.selectionStart : 0;
            var e = typeof ta.selectionEnd === "number" ? ta.selectionEnd : s;
            var up = norm(raw.slice(0, s));
            var curLine = (up.match(/\n/g) || []).length;
            var colInLine = up.length - (up.lastIndexOf("\n") + 1);
            return { raw: raw, s: s, e: e, curLine: curLine, colInLine: colInLine };
        }

        function canEnterAt(lineIdx) {
            return (lineIdx < maxLines - 1) && tailHasFree(lineIdx);
        }

        if (ta.getAttribute("data-eb-bind") === "1") return;
        ta.setAttribute("data-eb-bind", "1");

        ta.addEventListener("beforeinput", function (ev) {
            var type = ev.inputType || "";
            if (type.indexOf("delete") === 0) return;
            if (type === "insertFromPaste") return;

            var isEnter = (type === "insertParagraph" || type === "insertLineBreak");
            var data = isEnter ? "\n" : (ev.data || "");
            if (!data && !/insert|composition/i.test(type)) return;

            var meta = caretMeta();
            var raw = meta.raw, s = meta.s, e = meta.e, curLine = meta.curLine, colInLine = meta.colInLine;

            if (isEnter) {
                ev.preventDefault();
                if (!canEnterAt(curLine)) { blink(ta); return; }
                ta.value = raw.slice(0, s) + "\n" + raw.slice(e);
                var pos = s + 1;
                ta.setSelectionRange(pos, pos);
                window.queueMicrotask ? queueMicrotask(function () { __eb_normalizeClamp(ta, maxLines); }) :
                    setTimeout(function () { __eb_normalizeClamp(ta, maxLines); }, 0);
                return;
            }

            var lines = norm(raw).split("\n");
            var curLineText = lines[curLine] || "";

            var canvas = document.createElement("canvas");
            var ctx = canvas.getContext("2d");
            var cs = window.getComputedStyle(ta);
            ctx.font = [cs.fontStyle, cs.fontVariant, cs.fontWeight, cs.fontSize, cs.fontFamily].join(" ");

            var before = curLineText.slice(0, colInLine);
            var after = curLineText.slice(colInLine + (e - s));
            var candidate = before + (ev.data || "") + after;

            var tooWide = ctx.measureText(candidate).width > __eb_getAvailWidth(ta);
            if (tooWide) {
                if (!canEnterAt(curLine)) {
                    ev.preventDefault();
                    blink(ta);
                    return;
                }
                ev.preventDefault();
                ta.value = raw.slice(0, s) + "\n" + (ev.data || "") + raw.slice(e);
                var pos2 = s + 1 + (ev.data || "").length;
                ta.setSelectionRange(pos2, pos2);
                window.queueMicrotask ? queueMicrotask(function () { __eb_normalizeClamp(ta, maxLines); }) :
                    setTimeout(function () { __eb_normalizeClamp(ta, maxLines); }, 0);
                return;
            }

            var simulated = norm(raw.slice(0, s) + (ev.data || "") + raw.slice(e));
            var lineCount = (simulated.match(/\n/g) || []).length + 1;
            if (lineCount > maxLines) {
                ev.preventDefault();
                blink(ta);
            }
        }, true);

        ta.addEventListener("keydown", function (e) {
            if (e.key !== "Enter") return;
            var raw = ta.value;
            var s = typeof ta.selectionStart === "number" ? ta.selectionStart : 0;
            var up = raw.slice(0, s).replace(/\r\n?/g, "\n");
            var curLine = (up.match(/\n/g) || []).length;
            if (!canEnterAt(curLine)) {
                e.preventDefault();
                blink(ta);
            }
        }, true);
    }

    var TAB_SIZE = 4;
    var TAB_SPACES = (new Array(TAB_SIZE + 1)).join(" ");

    function __eb_insertSmart_fromStart(ta, text, maxLines) {
        function norm(s) { return String(s || "").replace(/\r\n?/g, "\n"); }

        var raw = norm(ta.value);
        var s = typeof ta.selectionStart === "number" ? ta.selectionStart : 0;
        var startLine = (norm(raw.slice(0, s)).match(/\n/g) || []).length;

        var parts = raw.split("\n").slice(0, startLine);
        ta.value = parts.join("\n");
        if (ta.value && ta.value.slice(-1) !== "\n") ta.value += "\n";
        ta.setSelectionRange(ta.value.length, ta.value.length);

        text = norm(text).replace(/\t/g, TAB_SPACES);

        var availW = __eb_getAvailWidth(ta);
        __eb_ctx.font = __eb_ctxFontFrom(ta);

        var tokens = [];
        var i, j, ch;
        for (i = 0; i < text.length;) {
            ch = text[i];
            if (ch === "\n") { tokens.push("\n"); i++; continue; }
            if (ch === " ") {
                j = i + 1;
                while (j < text.length && text[j] === " ") j++;
                tokens.push(text.slice(i, j));
                i = j; continue;
            }
            j = i + 1;
            while (j < text.length && text[j] !== "\n" && text[j] !== " ") j++;
            tokens.push(text.slice(i, j));
            i = j;
        }

        var lines = [""];
        var li = 0;

        function push(tok) {
            if (tok === "\n") {
                if (lines.length >= maxLines) return false;
                lines.push("");
                li++;
                return true;
            }
            var cur = lines[li] || "";
            var cand = cur + tok;
            if (__eb_ctx.measureText(cand).width <= availW) {
                lines[li] = cand;
                return true;
            }
            if (lines.length >= maxLines) return false;

            lines.push("");
            li++;
            if (__eb_ctx.measureText(tok).width <= availW) {
                lines[li] = tok;
                return true;
            }
            var k, c, t;
            for (k = 0; k < tok.length; k++) {
                c = tok.charAt(k);
                t = (lines[li] || "") + c;
                if (__eb_ctx.measureText(t).width > availW) {
                    if (lines.length >= maxLines) return false;
                    lines.push("");
                    li++;
                    if (__eb_ctx.measureText(c).width > availW) return false;
                    lines[li] = c;
                } else {
                    lines[li] = t;
                }
            }
            return true;
        }

        for (i = 0; i < tokens.length; i++) {
            if (!push(tokens[i])) break;
        }

        lines = lines.slice(0, maxLines);
        var body = lines.join("\n");
        var glue = (ta.value && body) ? "\n" : "";
        ta.value = ta.value + glue + body;
        __eb_normalizeClamp(ta, maxLines);
    }

    function bindTabAsSpaces(ta) {
        ta.addEventListener("keydown", function (e) {
            if (e.key !== "Tab") return;
            e.preventDefault();
            var s = typeof ta.selectionStart === "number" ? ta.selectionStart : 0;
            var ePos = typeof ta.selectionEnd === "number" ? ta.selectionEnd : s;
            var before = ta.value.slice(0, s);
            var after = ta.value.slice(ePos);
            var spaces = TAB_SPACES;
            ta.value = before + spaces + after;
            var pos = s + spaces.length;
            ta.setSelectionRange(pos, pos);
            ta.dispatchEvent(new Event("input"));
        });
    }

    function yToLineIdx(relY, lineH, maxLines) {
        var h = Math.max(1, lineH | 0);
        var idx = Math.floor(Math.max(0, Math.min(relY, maxLines * h - 1)) / h);
        if (idx < 0) idx = 0;
        if (idx > maxLines - 1) idx = maxLines - 1;
        return idx;
    }

    function moveCaretToLineEnd(ta, lineIdx) {
        var lines = ta.value.replace(/\r\n?/g, "\n").split("\n");
        var i = Math.max(0, Math.min(lineIdx, Math.max(0, lines.length - 1)));
        var pos = 0;
        var k;
        for (k = 0; k < i; k++) pos += (lines[k] || "").length + 1;
        pos += (lines[i] || "").length;
        ta.setSelectionRange(pos, pos);
    }

    function createGroupsWithEditorOverlay(isWrite) {
        if (!isWrite) {
            for (var i = 0; i < blocks.length; i++) {
                if (blocks[i].block && blocks[i].block.parentNode) blocks[i].block.parentNode.removeChild(blocks[i].block);
                if (blocks[i].ta && blocks[i].ta.parentNode) blocks[i].ta.parentNode.removeChild(blocks[i].ta);
            }
            blocks = [];
            window.__eb_blocks__ = blocks;
            return;
        }

        var inputs = (descriptor && descriptor.inputs && descriptor.inputs.length) ? descriptor.inputs : [];
        var by = {};

        for (var idx = 0; idx < inputs.length; idx++) {
            var f = inputs[idx];
            if (!f || !f.key || !f.a1) continue;
            var type = String(f.type || "Text").toLowerCase();
            if (type !== "text") continue;

            var m = String(f.key).match(/^(.*)_(\d+)$/);
            var rc = a1ToRC(f.a1);
            if (!rc) continue;

            var base = m ? m[1] : f.key;
            var sig = base + ":" + rc.c;
            if (!by[sig]) by[sig] = [];
            by[sig].push({ key: f.key, rc: rc });
        }

        for (var b = 0; b < blocks.length; b++) {
            if (blocks[b].block && blocks[b].block.parentNode) blocks[b].block.parentNode.removeChild(blocks[b].block);
            if (blocks[b].ta && blocks[b].ta.parentNode) blocks[b].ta.parentNode.removeChild(blocks[b].ta);
        }
        blocks = [];

        var host = document.getElementById("xhost");
        if (!host) {
            window.__eb_blocks__ = blocks;
            return;
        }

        var sig;
        for (sig in by) {
            if (!by.hasOwnProperty(sig)) continue;
            var list = by[sig];
            list.sort(function (a, b2) { return a.rc.r - b2.rc.r; });

            var runs = [];
            var run = [];
            for (var ii = 0; ii < list.length; ii++) {
                if (ii === 0) {
                    run = [list[ii]];
                } else {
                    var prev = list[ii - 1].rc.r;
                    var cur = list[ii].rc.r;
                    if (cur === prev + 1) run.push(list[ii]);
                    else {
                        if (run.length) runs.push({ keys: run.map(function (x) { return x.key; }) });
                        run = [list[ii]];
                    }
                }
            }
            if (run.length) runs.push({ keys: run.map(function (x) { return x.key; }) });

            for (var rgi = 0; rgi < runs.length; rgi++) {
                var keys = runs[rgi].keys.slice();
                var filtered = [];
                for (var k2 = 0; k2 < keys.length; k2++) {
                    if (tdByKey(keys[k2])) filtered.push(keys[k2]);
                }
                keys = filtered;
                if (!keys.length) continue;

                for (var k3 = 0; k3 < keys.length; k3++) {
                    var td = tdByKey(keys[k3]);
                    if (td) td.classList.add("eb-group");
                }

                var firstCell = cellDivByKey(keys[0]);
                if (!firstCell) continue;

                var initLines = [];
                for (var k4 = 0; k4 < keys.length; k4++) {
                    var cdiv = cellDivByKey(keys[k4]);
                    initLines.push(cdiv ? (cdiv.textContent || "") : "");
                }
                for (var k5 = 0; k5 < keys.length; k5++) {
                    var cdiv2 = cellDivByKey(keys[k5]);
                    if (cdiv2) cdiv2.textContent = "";
                }

                var block = document.createElement("div");
                block.className = "eb-block";
                block.setAttribute("tabindex", "-1");

                var ta = document.createElement("textarea");
                ta.className = "eb-ta";
                ta.setAttribute("rows", String(keys.length));
                ta.setAttribute("wrap", "off");

                copyStyleFromFirstCell(block, firstCell);
                copyStyleFromFirstCell(ta, firstCell);

                var placeRect = (function (keysCopy, blockEl, taEl) {
                    return function () {
                        var rct = blockRectOfKeys(keysCopy) || { left: 0, top: 0, width: 0, height: 0 };
                        var cRect = clampToTable({
                            left: Math.floor(rct.left),
                            top: Math.floor(rct.top),
                            width: Math.ceil(rct.width),
                            height: Math.ceil(rct.height + 2)
                        });
                        blockEl.style.left = cRect.left + "px";
                        blockEl.style.top = cRect.top + "px";
                        blockEl.style.width = cRect.width + "px";
                        blockEl.style.height = cRect.height + "px";

                        taEl.style.left = cRect.left + "px";
                        taEl.style.top = cRect.top + "px";
                        taEl.style.width = cRect.width + "px";
                        taEl.style.height = cRect.height + "px";
                    };
                })(keys, block, ta);

                placeRect();

                var hasText = false;
                for (var il = 0; il < initLines.length; il++) {
                    if (initLines[il] && initLines[il].length) { hasText = true; break; }
                }
                ta.value = hasText ? initLines.join("\n") : "";

                var lineH = parseFloat(window.getComputedStyle(ta).lineHeight) || 16;
                __eb_bindBeforeInputWidthClamp(ta, keys.length);

                (function (keysCopy, taEl) {
                    function syncPayload() {
                        var vis = taEl.value.replace(/\r\n?/g, "\n").split("\n");
                        if (vis.length < keysCopy.length) {
                            while (vis.length < keysCopy.length) vis.push("");
                        }
                        vis = vis.slice(0, keysCopy.length);
                        for (var i = 0; i < keysCopy.length; i++) {
                            payloadInputs[keysCopy[i]] = vis[i] || "";
                        }
                    }
                    syncPayload();

                    taEl.addEventListener("paste", function (e) {
                        e.preventDefault();
                        var text = (e.clipboardData && e.clipboardData.getData("text/plain")) || "";
                        text = text.replace(/\u200B/g, "");
                        __eb_insertSmart_fromStart(taEl, text, keysCopy.length);
                        __eb_normalizeClamp(taEl, keysCopy.length);
                        syncPayload();
                    });

                    bindTabAsSpaces(taEl);
                    taEl.addEventListener("input", function () {
                        __eb_normalizeClamp(taEl, keysCopy.length);
                        syncPayload();
                    });
                    taEl.addEventListener("compositionend", function () {
                        taEl.dispatchEvent(new Event("input"));
                    });
                })(keys, ta);

                (function (keysCopy, taEl) {
                    function onCellMouseDown(e) {
                        e.preventDefault();
                        taEl.focus({ preventScroll: true });
                        var rect = taEl.getBoundingClientRect();
                        var relY = Math.max(0, Math.min(rect.height - 1, e.clientY - rect.top));
                        var li = yToLineIdx(relY, lineH, keysCopy.length);
                        moveCaretToLineEnd(taEl, li);
                    }
                    for (var ki = 0; ki < keysCopy.length; ki++) {
                        var cellTd = tdByKey(keysCopy[ki]);
                        if (cellTd) cellTd.addEventListener("mousedown", onCellMouseDown);
                    }
                })(keys, ta);

                (function (keysCopy, blockEl, taEl) {
                    taEl.addEventListener("focus", function () {
                        var rct = blockRectOfKeys(keysCopy);
                        if (!rct) return;
                        blockEl.setAttribute("data-active", "1");
                        var cRect = clampToTable({
                            left: Math.floor(rct.left),
                            top: Math.floor(rct.top),
                            width: Math.ceil(rct.width),
                            height: Math.ceil(rct.height)
                        });
                        showRingForRect(cRect);
                    }, true);
                    taEl.addEventListener("blur", function () {
                        blockEl.setAttribute("data-active", "0");
                        hideRingIfIdle();
                    }, true);
                })(keys, block, ta);

                host.appendChild(block);
                host.appendChild(ta);

                blocks.push({
                    block: block,
                    ta: ta,
                    keys: keys,
                    placeRect: placeRect,
                    refreshRing: (function (keysCopy) {
                        return function () {
                            if (block.getAttribute("data-active") !== "1") return;
                            var rct = blockRectOfKeys(keysCopy);
                            if (!rct) return;
                            var cRect = clampToTable({
                                left: Math.floor(rct.left),
                                top: Math.floor(rct.top),
                                width: Math.ceil(rct.width),
                                height: Math.ceil(rct.height)
                            });
                            showRingForRect(cRect);
                        };
                    })(keys)
                });
            }
        }

        window.__eb_blocks__ = blocks;
        window.dispatchEvent(new CustomEvent("eb-editor-mounted"));

        window.addEventListener("resize", function () {
            for (var i = 0; i < blocks.length; i++) {
                blocks[i].placeRect();
                blocks[i].refreshRing();
            }
        });
    }

    /* ================= 프리뷰 마운트 ================= */

    function mount(hostSel, p, options) {
        options = options || {};
        var isWrite = (typeof options.isWrite === "boolean") ? options.isWrite : __EB_IS_WRITE__;
        var xhost = document.getElementById("xhost");
        if (!xhost || !p || !p.cells || !p.cells.length) {
            if (xhost) xhost.innerHTML = '<div class="alert alert-danger">DOC_Err_PreviewFailed</div>';
            return;
        }

        var xlPrev = document.getElementById("xlPreview");
        if (xlPrev) {
            xlPrev.style.fontFamily = "var(--eb-font)";
            xlPrev.style.letterSpacing = "normal";
            xlPrev.style.wordSpacing = "normal";
        }

        var styles = p.styles || {};
        var allRows = p.cells.length;
        var minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;

        function mark(r, c) {
            if (r < minR) minR = r;
            if (r > maxR) maxR = r;
            if (c < minC) minC = c;
            if (c > maxC) maxC = c;
        }

        var r, c;
        for (r = 1; r <= allRows; r++) {
            var row = p.cells[r - 1] || [];
            for (c = 1; c <= row.length; c++) {
                var v = row[c - 1];
                if (v !== "" && v != null) mark(r, c);
                var st = styles[r + "," + c];
                if (hasVisibleStyle(st)) mark(r, c);
            }
        }

        var merges = p.merges || [];
        for (var mi = 0; mi < merges.length; mi++) {
            var m = merges[mi];
            var r1 = parseInt(m[0], 10) || 0;
            var c1 = parseInt(m[1], 10) || 0;
            var r2 = parseInt(m[2], 10) || 0;
            var c2 = parseInt(m[3], 10) || 0;
            mark(r1, c1); mark(r2, c2);
        }

        var inInputs = (descriptor && descriptor.inputs && descriptor.inputs.length) ? descriptor.inputs : [];
        for (var ii = 0; ii < inInputs.length; ii++) {
            var f = inInputs[ii];
            var rc = a1ToRC(f && f.a1);
            if (rc) mark(rc.r, rc.c);
        }

        if (!isFinite(minR) || !isFinite(minC)) {
            minR = 1; maxR = 1; minC = 1; maxC = 1;
        }

        var maxColsFromCells = 1;
        for (r = 0; r < p.cells.length; r++) {
            if (p.cells[r].length > maxColsFromCells) maxColsFromCells = p.cells[r].length;
        }
        if (maxC > maxColsFromCells) maxC = maxColsFromCells;
        if (minC < 1) minC = 1;

        // 머지맵
        var mergeMap = new (window.Map || function () { this._ = {}; this.set = function (k, v) { this._[k] = v; }; this.get = function (k) { return this._[k]; }; })();
        for (mi = 0; mi < merges.length; mi++) {
            m = merges[mi];
            var rr1 = Math.max(parseInt(m[0], 10) || 0, minR);
            var cc1 = Math.max(parseInt(m[1], 10) || 0, minC);
            var rr2 = Math.min(parseInt(m[2], 10) || 0, maxR);
            var cc2 = Math.min(parseInt(m[3], 10) || 0, maxC);
            if (rr1 > rr2 || cc1 > cc2) continue;
            var masterKey = rr1 + "-" + cc1;
            mergeMap.set(masterKey, { master: true, rs: rr2 - rr1 + 1, cs: cc2 - cc1 + 1 });
            for (r = rr1; r <= rr2; r++) {
                for (c = cc1; c <= cc2; c++) {
                    var k = r + "-" + c;
                    if (k !== masterKey) mergeMap.set(k, { covered: true });
                }
            }
        }

        var styleGrid = [];
        for (r = 0; r <= maxR; r++) {
            styleGrid[r] = [];
            for (c = 0; c <= maxC; c++) styleGrid[r][c] = null;
        }

        for (r = minR; r <= maxR; r++) {
            for (c = minC; c <= maxC; c++) {
                var key = r + "," + c;
                var stRaw = styles[key] || {};
                var border = stRaw.border || {};
                styleGrid[r][c] = {
                    font: stRaw.font || null,
                    align: stRaw.align || null,
                    fill: stRaw.fill || null,
                    border: {
                        l: border.l || "None",
                        r: border.r || "None",
                        t: border.t || "None",
                        b: border.b || "None"
                    }
                };
            }
        }

        function weight(s) {
            s = String(s || "").toLowerCase();
            if (!s || s === "none") return 0;
            if (s.indexOf("double") >= 0) return 6;
            if (s.indexOf("thick") >= 0) return 5;
            if (s.indexOf("mediumdashdotdot") >= 0 ||
                s.indexOf("mediumdashdot") >= 0 ||
                s.indexOf("mediumdashed") >= 0 ||
                s.indexOf("medium") >= 0) return 4;
            if (s.indexOf("dashed") >= 0 || s.indexOf("dashdot") >= 0 || s.indexOf("dashdotdot") >= 0) return 3;
            if (s.indexOf("dotted") >= 0 || s.indexOf("hair") >= 0) return 2;
            return 1;
        }
        function stronger(a, b) {
            return (weight(a) >= weight(b)) ? a : b;
        }

        for (r = minR; r <= maxR; r++) {
            for (c = minC; c <= maxC; c++) {
                var cur = styleGrid[r][c];
                if (!cur) continue;
                if (c < maxC) {
                    var right = styleGrid[r][c + 1];
                    if (right) {
                        var pick = stronger(cur.border.r, right.border.l);
                        cur.border.r = pick;
                        right.border.l = pick;
                    }
                }
                if (r < maxR) {
                    var down = styleGrid[r + 1][c];
                    if (down) {
                        var pick2 = stronger(cur.border.b, down.border.t);
                        cur.border.b = pick2;
                        down.border.t = pick2;
                    }
                }
            }
        }

        function colPxAt(cidx) {
            var arr = p.colW || [];
            var wChar = arr[cidx - 1];
            return excelColWidthToPx(typeof wChar === "undefined" ? 8.43 : wChar);
        }
        function sumColPx(c1, c2) {
            var s = 0;
            for (var i = c1; i <= c2; i++) s += colPxAt(i);
            return s;
        }

        var tbl = document.createElement("table");
        tbl.className = "xlfb";

        var colgroup = document.createElement("colgroup");
        for (c = minC; c <= maxC; c++) {
            var cg = document.createElement("col");
            cg.style.width = colPxAt(c).toFixed(2) + "px";
            colgroup.appendChild(cg);
        }
        tbl.appendChild(colgroup);

        var tbody = document.createElement("tbody");
        var rowHeights = p.rowH || [];
        var DEFAULT_ROW_PT = 15;

        for (r = minR; r <= maxR; r++) {
            var tr = document.createElement("tr");
            var pt = (typeof rowHeights[r - 1] !== "undefined") ? rowHeights[r - 1] : DEFAULT_ROW_PT;
            var rowPx = ptToPx(pt);
            tr.style.height = rowPx + "px";

            for (c = minC; c <= maxC; c++) {
                var key2 = r + "-" + c;
                var mm = mergeMap.get ? mergeMap.get(key2) : mergeMap._ && mergeMap._[key2];
                if (mm && mm.covered) continue;

                var td = document.createElement("td");
                td.setAttribute("data-rowpx", String(rowPx));
                td.setAttribute("data-r", String(r));
                td.setAttribute("data-c", String(c));

                if (mm && mm.master) {
                    if (mm.rs > 1) td.setAttribute("rowspan", String(mm.rs));
                    if (mm.cs > 1) td.setAttribute("colspan", String(mm.cs));
                }
                if (!(mm && mm.master && mm.cs > 1)) {
                    td.style.width = colPxAt(c) + "px";
                }

                var cellDiv = document.createElement("div");
                cellDiv.className = "cellc";
                cellDiv.style.lineHeight = rowPx + "px";
                if (!mm) cellDiv.style.maxHeight = rowPx + "px";

                var v2 = (preview && preview.cells && preview.cells[r - 1] && typeof preview.cells[r - 1][c - 1] !== "undefined")
                    ? preview.cells[r - 1][c - 1] : "";
                if (v2 !== "") cellDiv.appendChild(document.createTextNode(String(v2)));

                applyStyleToCell(td, cellDiv, styleGrid[r][c]);

                var meta = posToMeta && posToMeta.get ? posToMeta.get(r + "," + c) : null;
                var fieldKey = meta && meta.key;
                var fieldType = (meta && meta.type ? meta.type : "Text").toLowerCase();
                var editable = !!fieldKey && !(mm && !mm.master) && isWrite;

                if (editable) {
                    td.setAttribute("data-key", fieldKey);
                    td.classList.add("eb-editable");

                    if (fieldType === "date") {
                        td.setAttribute("data-type", "date");
                        td.classList.add("eb-group");

                        var input = document.createElement("input");
                        input.type = "date";
                        input.className = "eb-input-date";

                        var today = new Date();
                        var yyyy = today.getFullYear();
                        var mm2 = ("0" + (today.getMonth() + 1)).slice(-2);
                        var dd = ("0" + today.getDate()).slice(-2);
                        input.value = yyyy + "-" + mm2 + "-" + dd;

                        cellDiv.textContent = "";
                        cellDiv.appendChild(input);

                        (function (fieldKeyCopy, inputEl) {
                            function sync() {
                                payloadInputs[fieldKeyCopy] = inputEl.value || "";
                            }
                            inputEl.addEventListener("change", sync);
                            inputEl.addEventListener("blur", sync);
                            if (!(fieldKeyCopy in payloadInputs)) sync();

                            td.addEventListener("mousedown", function (e) {
                                if (e.button !== 0) return;
                                e.preventDefault();
                                inputEl.focus({ preventScroll: true });
                                if (typeof inputEl.showPicker === "function") {
                                    try { inputEl.showPicker(); } catch (e2) { }
                                }
                            });
                        })(fieldKey, input);
                    }
                }

                td.appendChild(cellDiv);
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
        tbl.appendChild(tbody);

        var totalW = sumColPx(minC, maxC);
        tbl.style.width = totalW + "px";

        var xhostEl = document.getElementById("xhost");
        xhostEl.innerHTML = "";
        xhostEl.appendChild(tbl);

        window.__DOC_TOTALW__ = Math.max(1, totalW | 0);
        window.__DOC_EFFW__ = Math.max(1, measureEffectiveContentWidth(tbl) | 0);

        applyFinalWidth();
        window.requestAnimationFrame ?
            window.requestAnimationFrame(function () { updateClampBounds(); }) :
            setTimeout(updateClampBounds, 0);

        try {
            if (typeof window.colorizeInputGroups === "function") window.colorizeInputGroups(isWrite);
        } catch (e) { }

        createGroupsWithEditorOverlay(isWrite);
    }

    /* ================= 초기화 / 레이아웃 ================= */

    (function initHost() {
        var host = document.getElementById("doc-scroll");
        if (!host) return;
        escapeFixedTraps(host);
        placeContainer();
    })();

    function rerenderAll() {
        placeContainer();
        applyFinalWidth();
        updateClampBounds();
        if (window.__eb_blocks__) {
            var arr = window.__eb_blocks__;
            for (var i = 0; i < arr.length; i++) {
                if (arr[i].placeRect) arr[i].placeRect();
                if (arr[i].refreshRing) arr[i].refreshRing();
            }
        }
    }

    function firstReflow() {
        try {
            if (document.fonts && document.fonts.ready && document.fonts.ready.then) {
                document.fonts.ready.then(function () {
                    setTimeout(rerenderAll, 0);
                });
            } else {
                setTimeout(rerenderAll, 0);
            }
        } catch (e) {
            setTimeout(rerenderAll, 0);
        }
    }

    window.addEventListener("load", firstReflow, { once: true });
    firstReflow();
    window.addEventListener("resize", rerenderAll);

    var xhostObsTarget = document.getElementById("xhost");
    if (xhostObsTarget && window.MutationObserver) {
        var moQueued = false;
        var mo = new MutationObserver(function () {
            if (moQueued) return;
            moQueued = true;
            (window.requestAnimationFrame || setTimeout)(function () {
                var tbl = document.querySelector("#xhost table");
                if (tbl) {
                    window.__DOC_EFFW__ = Math.max(Number(window.__DOC_EFFW__ || 0) || 0, measureEffectiveContentWidth(tbl) | 0);
                }
                rerenderAll();
                moQueued = false;
            }, 0);
        });
        mo.observe(xhostObsTarget, { attributes: true, childList: true, subtree: true });
    }

    /* ================= 전역 API ================= */

    window.EBDocPreview = {
        T: T,
        info: info,
        ok: ok,
        err: err,
        descriptor: descriptor,
        preview: preview,
        get payloadInputs() { return payloadInputs; },
        mount: mount,
        rerenderAll: rerenderAll,
        firstReflow: firstReflow,
        escapeFixedTraps: escapeFixedTraps,
        placeContainer: placeContainer,
        readJson: readJson,
        setWriteMode: function (flag) {
            __EB_IS_WRITE__ = !!flag;
            window.__DOC_IS_WRITE__ = __EB_IS_WRITE__;
        },
        get isWrite() { return __EB_IS_WRITE__; },
        getMetrics: function () {
            return {
                totalW: window.__DOC_TOTALW__ || 0,
                effW: window.__DOC_EFFW__ || 0,
                finalW: window.__DOC_FINALW__ || 0,
                hMax: HMAX
            };
        }
    };

    function debugHorizontalBounds(tag) {
        try {
            var host = document.getElementById("doc-scroll");
            var xhost = document.getElementById("xhost");
            if (!host || !xhost) {
                console.log("[DOC DEBUG]", tag, "host/xhost not found");
                return;
            }
            var table = xhost.querySelector("table");
            var scrollWidth = host.scrollWidth;
            var clientWidth = host.clientWidth;
            var maxScroll = scrollWidth - clientWidth;
            var xhostW = xhost.offsetWidth;
            var tableW = table ? table.offsetWidth : null;

            console.log("[DOC DEBUG]", tag, {
                DOC_HMAX: window.__DOC_HMAX__,
                HMAX: HMAX,
                totalW: window.__DOC_TOTALW__,
                effW: window.__DOC_EFFW__,
                finalW: window.__DOC_FINALW__,
                scrollWidth: scrollWidth,
                clientWidth: clientWidth,
                maxScroll: maxScroll,
                xhostW: xhostW,
                tableW: tableW
            });
        } catch (e) {
            console.log("[DOC DEBUG] error in debugHorizontalBounds", e);
        }
    }
})();
