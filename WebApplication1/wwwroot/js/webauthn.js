// --- base64url helpers ---
const b64uToBuf = (b64u) => {
    const pad = '==='.slice((b64u.length + 3) % 4);
    const b64 = (b64u.replace(/-/g, '+').replace(/_/g, '/')) + pad;
    const bin = atob(b64);
    const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    return buf.buffer;
};
const bufToB64u = (buf) => {
    const bin = String.fromCharCode(...new Uint8Array(buf));
    return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
};

// --- fetch helpers ---
async function postJson(url, obj) {
    const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify(obj ?? {})
    });
    if (!res.ok) throw new Error(`${url} ${res.status}`);
    return await res.json();
}

// --- Register (로그인된 상태에서 사용) ---
async function registerPasskey() {
    // 1) 서버에서 options
    const options = await postJson('/Passkey/BeginRegister', {});

    // 2) 브라우저가 요구하는 타입으로 변환
    options.challenge = b64uToBuf(options.challenge);
    options.user.id = b64uToBuf(options.user.id);
    if (options.excludeCredentials) {
        options.excludeCredentials = options.excludeCredentials.map(x => ({
            ...x, id: b64uToBuf(x.id)
        }));
    }

    // 3) create()
    const cred = await navigator.credentials.create({ publicKey: options });

    // 4) 서버로 attestation 전송
    const attResp = {
        id: cred.id,
        rawId: bufToB64u(cred.rawId),
        type: cred.type,
        response: {
            attestationObject: bufToB64u(cred.response.attestationObject),
            clientDataJSON: bufToB64u(cred.response.clientDataJSON)
        }
    };
    await postJson('/Passkey/CompleteRegister', attResp);
    alert('Passkey 등록 완료');
}

// --- Login (2FA/패스워드리스용) ---
async function loginPasskey() {
    // 1) 서버에서 options
    const options = await postJson('/Passkey/BeginLogin', {});

    // 2) 변환
    options.challenge = b64uToBuf(options.challenge);
    if (options.allowCredentials) {
        options.allowCredentials = options.allowCredentials.map(x => ({
            ...x, id: b64uToBuf(x.id)
        }));
    }

    // 3) get()
    const assertion = await navigator.credentials.get({ publicKey: options });

    // 4) 서버로 assertion 전송
    const asResp = {
        id: assertion.id,
        rawId: bufToB64u(assertion.rawId),
        type: assertion.type,
        response: {
            authenticatorData: bufToB64u(assertion.response.authenticatorData),
            clientDataJSON: bufToB64u(assertion.response.clientDataJSON),
            signature: bufToB64u(assertion.response.signature),
            userHandle: assertion.response.userHandle ? bufToB64u(assertion.response.userHandle) : null
        }
    };
    await postJson('/Passkey/CompleteLogin', asResp);
    alert('Passkey 로그인 OK');
}

// --- 버튼 바인딩 ---
document.getElementById('btn-register')?.addEventListener('click', async () => {
    if (!('credentials' in navigator)) return alert('WebAuthn 미지원 브라우저');
    await registerPasskey().catch(e => { console.error(e); alert(e.message); });
});
document.getElementById('btn-login')?.addEventListener('click', async () => {
    if (!('credentials' in navigator)) return alert('WebAuthn 미지원 브라우저');
    await loginPasskey().catch(e => { console.error(e); alert(e.message); });
});
