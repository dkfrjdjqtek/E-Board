/* 2025.09.16 Added: EBValidate 기본 유틸 공개 */
// window.EBValidate 전역으로 노출
window.EBValidate = (function () {
    function el(x) { return (typeof x === 'string') ? document.getElementById(x) : x; }

    // --- 상단 경고 박스 ---
    function showAlert(container, messages) {
        const box = el(container);
        if (!box) return;
        const list = Array.isArray(messages) ? messages : (messages ? [messages] : []);
        let ul = box.querySelector('ul');
        if (!ul) {
            ul = document.createElement('ul');
            ul.className = 'mb-0';
            box.appendChild(ul);
        }
        ul.innerHTML = '';
        list.filter(Boolean).forEach(m => {
            const li = document.createElement('li');
            li.textContent = m;
            ul.appendChild(li);
        });
        box.classList.toggle('d-none', list.length === 0);
    }
    function clearAlert(container) { showAlert(container, []); }

    // --- 필드 단위 에러 표시 ---
    function findFieldBox(input) {
        return input.closest(
            '.mb-3, .col, .col-12, .col-md-4, .col-md-6, .form-group, .form-floating, .input-group'
        ) || input.parentElement;
    }

    function ensureFeedback(input, useTooltip) {
        const box = findFieldBox(input);
        const klass = useTooltip ? 'invalid-tooltip' : 'invalid-feedback';
        // 입력별로 고유 feedback 재사용
        const sel = `.${klass}[data-for="${input.id}"]`;
        let fb = box.querySelector(sel);
        if (!fb) {
            fb = document.createElement('div');
            fb.className = klass + (useTooltip ? '' : ' d-block');
            if (input.id) fb.setAttribute('data-for', input.id);
            box.appendChild(fb);
            if (useTooltip) box.classList.add('position-relative');
        }
        return fb;
    }

    // 기본: 아래 문구형. 툴팁으로 쓰려면 3번째 인자 true
    function setInvalid(input, message, useTooltip = false) {
        if (!input) return;
        input.classList.add('is-invalid');
        const fb = ensureFeedback(input, useTooltip);
        fb.textContent = message || '';
    }

    function clearInvalid(input) {
        if (!input) return;
        input.classList.remove('is-invalid', 'is-valid');
        const box = findFieldBox(input);
        // 이 입력에 대응하는 피드백만 제거
        if (input.id) {
            box.querySelectorAll(`.invalid-feedback[data-for="${input.id}"], .invalid-tooltip[data-for="${input.id}"]`)
                .forEach(n => n.remove());
        } else {
            // id가 없으면 박스 내 전부 제거(예외 케이스)
            box.querySelectorAll('.invalid-feedback, .invalid-tooltip').forEach(n => n.remove());
        }
    }

    function clearAll(scope) {
        const root = el(scope) || document;
        root.querySelectorAll('.is-invalid, .is-valid').forEach(i => i.classList.remove('is-invalid', 'is-valid'));
        root.querySelectorAll('.invalid-feedback, .invalid-tooltip').forEach(n => n.remove());
    }

    // (옵션) 여러 필드를 한 번에 표시하고 싶을 때: { inputElOrId: "메시지", ... }
    function setErrorsMap(map, useTooltip = false) {
        if (!map) return;
        Object.entries(map).forEach(([k, msg]) => {
            const input = el(k);
            if (input) setInvalid(input, msg, useTooltip);
        });
    }

    /* ===================================================================== */
    /* 2025.09.17 Added: jQuery Validate 연동을 EBValidate 한 곳에서 고정   */
    /* - 제출 버튼 클릭 전까지 invalid 고정 유지                            */
    /* - onkeyup onfocusout onclick 비활성화로 포커스 이동시 상태 안 바뀜    */
    /* - invalid-form 시 _ValidationAlert와 필드 invalid를 EBValidate로 출력 */
    /* - 다음 제출 시 이전 락 해제 후 재계산                                 */
    /* ===================================================================== */

    // 2025.09.17 Added: 입력별 invalid 락 관리용
    const _lockObserverMap = new WeakMap(); // input -> MutationObserver

    function _lockInvalid(input, message) {
        // 2025.09.17 Added: 잠금 플래그와 즉시 invalid 반영
        input.setAttribute('data-ebv-lock', '1');                   // 2025.09.17 Added: 잠금 표식
        setInvalid(input, message || '');
        // 2025.09.17 Added: 누가 클래스를 지워도 재부여
        try {
            const obs = new MutationObserver(() => {
                if (input.getAttribute('data-ebv-lock') === '1' && !input.classList.contains('is-invalid')) {
                    input.classList.add('is-invalid');
                }
            });
            obs.observe(input, { attributes: true, attributeFilter: ['class'] });
            _lockObserverMap.set(input, obs);
        } catch { /* 2025.09.17 Added: 옵저버 미지원 브라우저 무시 */ }
    }

    function _unlockInvalid(input) {
        input.removeAttribute('data-ebv-lock');
        const obs = _lockObserverMap.get(input);
        if (obs) { try { obs.disconnect(); } catch { } }
        _lockObserverMap.delete(input);
        // 메시지/보더는 다음 검증 결과에 의해 재설정됨
        clearInvalid(input);
    }

    function _unlockAll(form) {
        (form.querySelectorAll('[data-ebv-lock="1"]') || []).forEach(_unlockInvalid);
    }

    // 2025.09.17 Added: 폼 바인딩 엔트리포인트
    function bindForm(formOrSel, opts) {
        const form = (typeof formOrSel === 'string') ? document.querySelector(formOrSel) : formOrSel;
        if (!form) return;
        const alertId = (opts && opts.alertId) || form.getAttribute('data-ebv-alert') || 'login-alert';

        // jQuery Validate 없으면 필수만 EBV로 표시
        if (!window.jQuery || !jQuery.validator) {
            form.addEventListener('submit', function (e) {
                const invalids = Array.from(form.querySelectorAll('[required]')).filter(i => !i.value);
                if (invalids.length) {
                    e.preventDefault();
                    const msgs = invalids.map(i => (i.getAttribute('data-val-required') || '필수입니다.'));
                    showAlert(alertId, msgs);
                    invalids.forEach(i => _lockInvalid(i, i.getAttribute('data-val-required') || '필수입니다.'));
                }
            });
            return;
        }

        // 2025.09.17 Changed: 전역 기본값을 EBV 정책으로 통일
        jQuery.validator.setDefaults({
            onkeyup: false,
            onfocusout: false,
            onclick: false,
            highlight: function (element) {
                if (element && element.classList) element.classList.add('is-invalid');
            },
            // 2025.09.17 Changed: 언하이라이트 금지 다음 제출 전까지 유지
            unhighlight: function () { /* no-op */ }
        });

        // 2025.09.17 Added: unobtrusive 파싱 보장 후 설정 적용
        try {
            jQuery(form).removeData('validator').removeData('unobtrusiveValidation');
            if (jQuery.validator.unobtrusive && jQuery.validator.unobtrusive.parse) {
                jQuery.validator.unobtrusive.parse(form);
            }
        } catch { /* ignore */ }

        const $f = jQuery(form);
        const v = $f.data('validator');
        if (v) {
            v.settings.onkeyup = false;
            v.settings.onfocusout = false;
            v.settings.onclick = false;
            // submitHandler는 통과 시 클린업
            const prevSubmit = v.settings.submitHandler;
            v.settings.submitHandler = function (formEl) {
                _unlockAll(form);               // 2025.09.17 Added: 성공 제출 전 락 해제
                if (typeof prevSubmit === 'function') return prevSubmit(formEl);
                formEl.submit();
            };
        }

        // 2025.09.17 Added: 제출 실패 시 한 번만 요약과 invalid를 고정 출력
        $f.on('invalid-form.validate', function (e, validator) {
            const list = (validator && Array.isArray(validator.errorList)) ? validator.errorList : [];
            const msgs = list.map(it => it && it.message).filter(Boolean);
            showAlert(alertId, msgs);
            // 이전 락 제거 후 현재 에러로 재설정
            _unlockAll(form);
            list.forEach(it => {
                if (it && it.element) _lockInvalid(it.element, it.message || '');
            });
        });

        // 2025.09.17 Added: form reset 시 전체 정리
        form.addEventListener('reset', function () {
            clearAlert(alertId);
            _unlockAll(form);
        });
    }

    // 2025.09.17 Added: 자동 바인딩 헬퍼 data-ebv-bind="auto" 가 지정된 폼만 대상
    (function autoBindOnReady() {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', autoBindOnReady);
            return;
        }
        document.querySelectorAll('form[data-ebv-bind="auto"]').forEach(f => bindForm(f));
    })();

    // 2025.09.17 Changed: bindForm 공개하여 뷰에서 한 줄로 연결 가능
    return { showAlert, clearAlert, setInvalid, clearInvalid, clearAll, setErrorsMap, bindForm };
})();
