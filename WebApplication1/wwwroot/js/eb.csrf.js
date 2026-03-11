
// 2026.01.26 Changed: 최초 Csrf fetch 실패 후에도 재시도 가능하도록 ensureToken 기반으로 수정하고, 불필요한 중괄호로 인한 스크립트 파싱 오류를 제거함

(function (w) {
    'use strict';

    const state = {
        headerName: 'RequestVerificationToken',
        token: '',
        inflight: null
    };

    async function ensureToken() {
        if (state.token) return state.token;
        if (state.inflight) return state.inflight;

        state.inflight = fetch('/Doc/Csrf', {
            method: 'GET',
            credentials: 'same-origin',
            headers: { 'Accept': 'application/json' }
        })
            .then(async r => {
                if (!r.ok) {
                    const err = new Error('csrf request failed');
                    err.status = r.status;
                    err.url = r.url;
                    throw err;
                }
                return r.json();
            })
            .then(j => {
                state.headerName = (j && j.headerName) ? String(j.headerName) : 'RequestVerificationToken';
                state.token = (j && j.token) ? String(j.token) : '';
                return state.token;
            })
            .catch(err => {
                console.warn('[EBCSRF] token fetch failed', err);
                return '';
            })
            .finally(() => {
                state.inflight = null;
            });

        return state.inflight;
    }

    function headers(extra) {
        const h = new Headers(extra || {});
        if (state.token) h.set(state.headerName, state.token);
        return h;
    }

    async function fetchWithCsrf(input, init) {
        const opts = Object.assign({}, init || {});
        const method = (opts.method || 'GET').toUpperCase();
        const needsToken = !['GET', 'HEAD', 'OPTIONS'].includes(method);

        if (!opts.credentials) opts.credentials = 'same-origin';

        const hasBody = opts.body != null;
        if (hasBody && !(opts.body instanceof FormData) && typeof opts.body !== 'string') {
            opts.headers = new Headers(opts.headers || {});
            if (!opts.headers.has('Content-Type')) {
                opts.headers.set('Content-Type', 'application/json');
            }
            opts.body = JSON.stringify(opts.body);
        }

        if (needsToken) {
            await ensureToken();
            const merged = new Headers(opts.headers || {});
            if (state.token) merged.set(state.headerName, state.token);
            opts.headers = merged;
        }

        return fetch(input, opts);
    }

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

    async function ensureFormHidden() {
        await ensureToken();
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

    if (!w.EBCSRF) {
        w.EBCSRF = {
            fetch: fetchWithCsrf,
            json: jsonWithCsrf,
            headers,
            ready: Promise.resolve(),
            headerName: () => state.headerName,
            token: () => state.token
        };
    }

    if (!w.EBUpload) {
        w.EBUpload = {};
    }

    const DEFAULTS = {
        url: '/Doc/Upload',
        maxBytes: 20 * 1024 * 1024,
        allowExt: [],
        allowMimePrefix: [],
        t: (k) => k
    };

    function extOf(name) {
        const m = (name || '').match(/\.([^.]+)$/);
        return m ? m[1].toLowerCase() : '';
    }

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

    async function uploadFiles(fileList, options) {
        const opt = Object.assign({}, DEFAULTS, options || {});
        const arr = Array.from(fileList || []).filter(Boolean);

        const vErrs = validateFiles(arr, opt);
        if (vErrs.length) {
            const err = new Error('validation failed');
            err.payload = { messages: vErrs, fieldErrors: null };
            throw err;
        }

        const fd = new FormData();
        for (const f of arr) fd.append('files', f, f.name);

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

        const items = Array.isArray(data.items) ? data.items : [];
        return items.map(x => ({
            FileKey: x.FileKey || '',
            OriginalName: x.OriginalName || '',
            ContentType: x.ContentType || '',
            ByteSize: typeof x.ByteSize === 'number' ? x.ByteSize : 0
        }));
    }

    w.EBUpload.uploadFiles = uploadFiles;
})(window);
