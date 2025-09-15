/* wwwroot/js/eb_validate.js */
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

    return { showAlert, clearAlert, setInvalid, clearInvalid, clearAll, setErrorsMap };
})();
