// 2025.10.15 Changed: 중복 정의 제거 단일 모듈로 통합 fetch 한번만 호출하도록 정리
// 2025.10.15 Added: EBCSRF.fetch EBCSRF.json EBCSRF.headers EBCSRF.ready 공개 API 일원화
// 2025.10.15 Added: POST PUT PATCH DELETE 시에만 자동 토큰 부착 FormData 전송 시 Content Type 미설정 유지
// 2025.10.15 Added: 폼 제출 시 __RequestVerificationToken 히든 필드 자동 주입

(function (w) {
    'use strict';

    // 2025.10.15 Added: 내부 상태
    const state = {
        headerName: 'RequestVerificationToken',
        token: '',
        inflight: null,
        ready: null
    };

    // 2025.10.15 Added: 토큰 확보
    async function ensureToken() {
        if (state.token) return state.token;
        if (state.inflight) return state.inflight;

        state.inflight = fetch('/Doc/Csrf', { method: 'GET', credentials: 'same-origin' })
            .then(r => {
                if (!r.ok) throw new Error('csrf request failed');
                return r.json();
            })
            .then(j => {
                state.headerName = (j && j.headerName) ? String(j.headerName) : 'RequestVerificationToken';
                state.token = (j && j.token) ? String(j.token) : '';
                state.inflight = null;
                return state.token;
            })
            .catch(err => {
                state.inflight = null;
                // 2025.10.15 Added: 토큰이 없어도 API는 동작하도록 유지
                console.warn('[EBCSRF] token fetch failed', err);
                return '';
            });

        return state.inflight;
    }

    // 2025.10.15 Added: 최초 준비 프라미스 노출
    state.ready = (async () => { await ensureToken(); })();

    // 2025.10.15 Added: 헤더 병합 도우미
    function headers(extra) {
        const h = new Headers(extra || {});
        if (state.token) h.set(state.headerName, state.token);
        return h;
    }

    // 2025.10.15 Added: fetch 래퍼 메서드
    async function fetchWithCsrf(input, init) {
        const opts = Object.assign({}, init || {});
        const method = (opts.method || 'GET').toUpperCase();
        const needsToken = !['GET', 'HEAD', 'OPTIONS'].includes(method);

        // 2025.10.15 Added: 동일 출처 쿠키 포함
        if (!opts.credentials) opts.credentials = 'same-origin';

        // 2025.10.15 Added: JSON 바디 자동 처리 FormData는 건드리지 않음
        const hasBody = opts.body != null;
        if (hasBody && !(opts.body instanceof FormData) && typeof opts.body !== 'string') {
            opts.headers = new Headers(opts.headers || {});
            if (!opts.headers.has('Content-Type')) {
                opts.headers.set('Content-Type', 'application/json');
            }
            opts.body = JSON.stringify(opts.body);
        }

        // 2025.10.15 Added: 변경 메서드에만 토큰 부착
        if (needsToken) {
            await state.ready;
            const merged = new Headers(opts.headers || {});
            if (state.token) merged.set(state.headerName, state.token);
            opts.headers = merged;
        }

        return fetch(input, opts);
    }

    // 2025.10.15 Added: JSON 헬퍼
    async function jsonWithCsrf(url, payload, method) {
        const m = (method || 'POST').toUpperCase();
        const res = await fetchWithCsrf(url, { method: m, body: payload });
        let data = null;
        try { data = await res.json(); } catch { data = null; }

        if (!res.ok) {
            const err = new Error('request failed');
            err.payload = data || { messages: ['DOC_Err_RequestFailed'] };
            throw err;
        }
        return data;
    }

    // 2025.10.15 Added: 폼 자동 히든 주입
    async function ensureFormHidden() {
        await state.ready;
        if (!state.token) return;
        try {
            const forms = document.querySelectorAll("form[method='post'], form[method='POST']");
            forms.forEach(f => {
                if (f.querySelector("input[name='__RequestVerificationToken']")) return;
                const input = document.createElement('input');
                input.type = 'hidden';
                input.name = '__RequestVerificationToken';
                input.value = state.token;
                f.appendChild(input);
            });
        } catch { /* no-op */ }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', ensureFormHidden, { once: true });
    } else {
        ensureFormHidden();
    }

    // 2025.10.15 Changed: 전역 내보내기 단일 객체로 고정 중복 초기화 방지
    if (!w.EBCSRF) {
        w.EBCSRF = {
            fetch: fetchWithCsrf,
            json: jsonWithCsrf,
            headers,
            ready: state.ready,
            headerName: () => state.headerName,
            token: () => state.token
        };
    }
    if (!w.EBUpload) {
        w.EBUpload = {};
    }

    // 2025.10.15 Added: 기본 옵션
    const DEFAULTS = {
        url: '/Doc/Upload',
        // 바이트 기준 (예: 20MB)
        maxBytes: 20 * 1024 * 1024,
        // 허용 확장자(소문자, 점 제외). 빈 배열이면 전체 허용
        allowExt: [],
        // 허용 MIME prefix. 빈 배열이면 전체 허용
        allowMimePrefix: [],
        // 다국어 키 리턴용 콜백(없으면 그대로 키 반환)
        t: (k) => k
    };

    // 2025.10.15 Added: 확장자 추출
    function extOf(name) {
        const m = (name || '').match(/\.([^.]+)$/);
        return m ? m[1].toLowerCase() : '';
    }

    // 2025.10.15 Added: 클라이언트 검증
    function validateFiles(files, opt) {
        const errs = [];

        for (const f of files) {
            if (!f) continue;
            if (opt.maxBytes && f.size > opt.maxBytes) {
                errs.push(opt.t('DOC_File_Err_TooLarge') + `: ${f.name}`);
            }
            if (Array.isArray(opt.allowExt) && opt.allowExt.length > 0) {
                const ex = extOf(f.name);
                if (!opt.allowExt.includes(ex)) {
                    errs.push(opt.t('DOC_File_Err_ExtNotAllowed') + `: ${f.name}`);
                }
            }
            if (Array.isArray(opt.allowMimePrefix) && opt.allowMimePrefix.length > 0) {
                const ok = opt.allowMimePrefix.some(p => (f.type || '').toLowerCase().startsWith(String(p).toLowerCase()));
                if (!ok) {
                    errs.push(opt.t('DOC_File_Err_MimeNotAllowed') + `: ${f.name}`);
                }
            }
        }
        return errs;
    }

    // 2025.10.15 Added: 업로드 실행 (FormData, CSRF 자동)
    async function uploadFiles(fileList, options) {
        const opt = Object.assign({}, DEFAULTS, options || {});
        const arr = Array.from(fileList || []).filter(Boolean);

        // 사전 검증
        const vErrs = validateFiles(arr, opt);
        if (vErrs.length) {
            const err = new Error('validation failed');
            err.payload = { messages: vErrs, fieldErrors: null };
            throw err;
        }

        // FormData 구성 (input name: files)
        const fd = new FormData();
        for (const f of arr) fd.append('files', f, f.name);

        // 업로드 호출 (multipart/form-data는 브라우저가 Content-Type 설정)
        const res = await (w.EBCSRF?.fetch || fetch)(opt.url, {
            method: 'POST',
            body: fd,
            credentials: 'same-origin'
        });

        let data = null;
        try { data = await res.json(); } catch { data = null; }

        if (!res.ok || !data) {
            const payload = (data && (data.messages || data.fieldErrors)) ? data : { messages: [opt.t('DOC_File_Err_UploadFailed')] };
            const err = new Error('upload failed');
            err.payload = payload;
            throw err;
        }

        // 예상 형식: { ok: true, items: [ { fileKey, originalName, contentType, byteSize } ... ] }
        const items = Array.isArray(data.items) ? data.items : [];
        return items.map(x => ({
            fileKey: x.fileKey || '',
            originalName: x.originalName || '',
            contentType: x.contentType || '',
            byteSize: typeof x.byteSize === 'number' ? x.byteSize : 0
        }));
    }

    // 2025.10.15 Added: 전역 export
    w.EBUpload.uploadFiles = uploadFiles;
})(window);
