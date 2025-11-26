// 2025.11.20 Changed: 공통 프리뷰 테이블(.xlfb)을 그대로 사용하고, A1 클릭/드래그 선택만 담당하도록 단순화. 자체 미리보기 렌더링 코드 제거.

(function (global) {
    "use strict";

    function rcToA1(row, col) {
        // 1-based row/col → A1
        let c = col;
        let letters = "";
        while (c > 0) {
            const rem = (c - 1) % 26;
            letters = String.fromCharCode(65 + rem) + letters;
            c = Math.floor((c - 1) / 26);
        }
        return letters + String(row);
    }

    function buildCellModel(table) {
        // tbody 기준으로 RC 정보/DOM 보관
        const rows = Array.from(table.tBodies[0]?.rows || []);
        const cells = [];
        rows.forEach((tr, rIdx) => {
            const cols = Array.from(tr.cells || []);
            cols.forEach((td, cIdx) => {
                const row = rIdx + 1;
                const col = cIdx + 1;
                let a1 = td.dataset.a1;
                if (!a1) {
                    a1 = rcToA1(row, col);
                    td.dataset.a1 = a1;
                }
                cells.push({ td, row, col, a1 });
            });
        });
        return { rows, cells };
    }

    function bindSelection(table) {
        if (!table || table._ebMapBound) return;
        table._ebMapBound = true;

        const model = buildCellModel(table);
        if (!model.cells.length) {
            console.warn("[EBMapPreview] no cells to bind.");
            return;
        }

        let dragging = false;
        let start = null;
        let end = null;

        function clearSel() {
            table.querySelectorAll("td.sel").forEach(td => td.classList.remove("sel"));
        }

        function applySel() {
            clearSel();
            if (!start || !end) return;

            const r1 = Math.min(start.row, end.row);
            const r2 = Math.max(start.row, end.row);
            const c1 = Math.min(start.col, end.col);
            const c2 = Math.max(start.col, end.col);

            model.cells.forEach(c => {
                if (c.row >= r1 && c.row <= r2 && c.col >= c1 && c.col <= c2) {
                    c.td.classList.add("sel");
                }
            });
        }

        function finishSelection() {
            if (!start || !end) { dragging = false; return; }

            const r1 = Math.min(start.row, end.row);
            const r2 = Math.max(start.row, end.row);
            const c1 = Math.min(start.col, end.col);
            const c2 = Math.max(start.col, end.col);

            let a1;
            if (r1 === r2 && c1 === c2) {
                a1 = rcToA1(r1, c1); // 단일 셀
            } else if (c1 === c2) {
                a1 = rcToA1(r1, c1) + ":" + rcToA1(r2, c2); // 단일 열 범위
            } else {
                // 직사각형 범위 전체
                a1 = rcToA1(r1, c1) + ":" + rcToA1(r2, c2);
            }

            // 포커스된 A1 입력이 등록한 setter 호출
            try {
                if (global.__a1 && typeof global.__a1.set === "function") {
                    global.__a1.set(a1);
                } else {
                    console.warn("[EBMapPreview] __a1.set 가 정의되어 있지 않습니다. 선택값:", a1);
                }
            } catch (e) {
                console.error("[EBMapPreview] __a1.set 호출 중 오류", e);
            }

            dragging = false;
        }

        function findCell(target) {
            if (!(target instanceof HTMLElement)) return null;
            const td = target.closest("td");
            if (!td || !table.contains(td)) return null;
            const a1 = td.dataset.a1;
            if (!a1) return null;

            // dataset.a1 에서 row/col 역추출(안 되면 rcToA1와 동일 규칙 사용)
            const m = String(a1).toUpperCase().match(/^([A-Z]+)(\d+)$/);
            let row, col;
            if (m) {
                row = parseInt(m[2], 10);
                // 컬럼 문자열 → 번호
                const s = m[1];
                let n = 0;
                for (let i = 0; i < s.length; i++) {
                    n = n * 26 + (s.charCodeAt(i) - 64);
                }
                col = n;
            } else {
                // fallback: DOM index
                const tr = td.parentElement;
                row = Array.prototype.indexOf.call(table.tBodies[0].rows, tr) + 1;
                col = Array.prototype.indexOf.call(tr.cells, td) + 1;
            }
            return { td, row, col, a1: a1 || rcToA1(row, col) };
        }

        table.addEventListener("mousedown", function (ev) {
            if (ev.button !== 0) return; // 좌클릭만
            const hit = findCell(ev.target);
            if (!hit) return;

            ev.preventDefault();
            dragging = true;
            start = { row: hit.row, col: hit.col };
            end = { row: hit.row, col: hit.col };
            table.classList.add("dragging");
            applySel();
        });

        table.addEventListener("mousemove", function (ev) {
            if (!dragging) return;
            const hit = findCell(ev.target);
            if (!hit) return;
            end = { row: hit.row, col: hit.col };
            applySel();
        });

        document.addEventListener("mouseup", function () {
            if (!dragging) return;
            table.classList.remove("dragging");
            finishSelection();
        });

        table.addEventListener("mouseleave", function () {
            // 드래그 중 바깥으로 나가면 마우스업과 동일 처리
            if (!dragging) return;
            table.classList.remove("dragging");
            finishSelection();
        });

        console.log("[EBMapPreview] selection bind 완료:", model.cells.length, "cells");
    }

    function bind() {
        // 공통 프리뷰(eb.doc.preview.common.js)가 렌더링한 #xhost 안의 첫 번째 테이블을 사용
        const host = document.getElementById("xhost") || document.getElementById("xlPreview");
        if (!host) {
            console.warn("[EBMapPreview] #xhost / #xlPreview 를 찾을 수 없습니다.");
            return;
        }

        // mount 직후 DOM 생성이 약간 지연될 수 있으므로 한 번 더 지연 체크
        const tryBind = function () {
            const table = host.querySelector("table");
            if (!table) {
                console.warn("[EBMapPreview] preview table 이 아직 없습니다.");
                return;
            }
            bindSelection(table);
        };

        if (host.querySelector("table")) {
            tryBind();
        } else {
            // 렌더가 아주 조금 늦을 때 대비
            setTimeout(tryBind, 30);
        }
    }

    global.EBMapPreview = {
        bind
    };
})(window);
