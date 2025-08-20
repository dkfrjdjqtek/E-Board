// Controllers/PasskeyController.cs
using System;
using System.Linq;
using System.Text;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Fido2NetLib;
using Fido2NetLib.Objects;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;

[Authorize]
// [IgnoreAntiforgeryToken] // 개발 중 CSRF 임시 해제 시 주석 제거
public class PasskeyController : Controller
{
    private readonly IFido2 _fido2;
    private readonly UserManager<IdentityUser> _users;

    private static readonly Dictionary<string, CredentialCreateOptions> _regCache = new();
    private static readonly Dictionary<string, AssertionOptions> _asrtCache = new();

    private record StoredCred(byte[] CredentialId, byte[] PublicKey, uint SignCount, byte[] UserHandle, string UserId);
    private static readonly ConcurrentDictionary<string, StoredCred> _store = new();

    public PasskeyController(IFido2 fido2, UserManager<IdentityUser> users)
    {
        _fido2 = fido2;
        _users = users;
    }

    // ===== Register =====
    [HttpPost]
    public async Task<IActionResult> BeginRegister([FromBody] object _)
    {
        var user = await _users.GetUserAsync(User);
        if (user == null) return Unauthorized();

        var fidoUser = new Fido2User
        {
            Id = Encoding.UTF8.GetBytes(user.Id),
            Name = user.UserName ?? user.Id,
            DisplayName = user.UserName ?? user.Id
        };

        var existing = new List<PublicKeyCredentialDescriptor>();

        var regOpts = _fido2.RequestNewCredential(fidoUser, existing); // 3.x: 2인자
        regOpts.AuthenticatorSelection = new AuthenticatorSelection
        {
            UserVerification = UserVerificationRequirement.Preferred
        };
        regOpts.Attestation = AttestationConveyancePreference.None;

        _regCache[user.Id] = regOpts;
        return Json(regOpts);
    }
    [HttpPost]
    public async Task<IActionResult> CompleteRegister(
    [FromBody] AuthenticatorAttestationRawResponse att, CancellationToken ct)
    {
        var user = await _users.GetUserAsync(User);
        if (user == null) return Unauthorized();
        if (!_regCache.TryGetValue(user.Id, out var regOpts))
            return BadRequest("registration options not found");

        var regResult = await _fido2.MakeNewCredentialAsync(
    att,
    regOpts,
    (args, token) => Task.FromResult(!_store.Values.Any(c => c.CredentialId.SequenceEqual(args.CredentialId))));


        //var key = Convert.ToBase64String(regResult.Result.CredentialId);
        var key = Convert.ToBase64String(regResult.Result!.CredentialId);
        _store[key] = new StoredCred(
            regResult.Result.CredentialId,
            regResult.Result.PublicKey,   // byte[]
            regResult.Result.Counter,     // uint
            Encoding.UTF8.GetBytes(user.Id),
            user.Id
        );

        return Ok(new { ok = true });
    }


    // ===== Login =====
    [AllowAnonymous]
    [HttpPost]
    public async Task<IActionResult> BeginLogin([FromBody] object _)
    {
        var user = await _users.GetUserAsync(User);
        if (user == null) return Unauthorized();

        var allowed = _store.Values
            .Where(c => c.UserId == user.Id)
            .Select(c => new PublicKeyCredentialDescriptor(c.CredentialId))
            .ToList();

        var asrtOpts = _fido2.GetAssertionOptions(
            allowedCredentials: allowed,
            userVerification: UserVerificationRequirement.Preferred);

        _asrtCache[user.Id] = asrtOpts;
        return Json(asrtOpts);
    }

    [AllowAnonymous]
    [HttpPost]
    public async Task<IActionResult> CompleteLogin(
        [FromBody] AuthenticatorAssertionRawResponse assertion, CancellationToken ct)
    {
        var user = await _users.GetUserAsync(User);
        if (user == null) return Unauthorized();
        if (!_asrtCache.TryGetValue(user.Id, out var asrtOpts))
            return BadRequest("assertion options not found");

        var credId = assertion.Id ?? assertion.RawId;
        if (credId == null || credId.Length == 0) return BadRequest("credential id missing");
        var key = Convert.ToBase64String(credId);

        if (!_store.TryGetValue(key, out var cred))
            return BadRequest("unknown credential");

        IsUserHandleOwnerOfCredentialIdAsync isOwner = (args, token) =>
        {
            var ok = cred.UserHandle.SequenceEqual(args.UserHandle)
                  && cred.CredentialId.SequenceEqual(args.CredentialId);
            return Task.FromResult(ok);
        };

        byte[] storedPublicKey = cred.PublicKey ?? Array.Empty<byte>();
        uint storedCounter = cred.SignCount;
        IsUserHandleOwnerOfCredentialIdAsync owner = isOwner;
        CancellationToken token = ct;

        //var assertResult = await _fido2.MakeAssertionAsync(assertion, asrtOpts, cred.PublicKey, cred.SignCount, isOwner, cred.UserHandle);
        var assertResult = await _fido2.MakeAssertionAsync(assertion, asrtOpts, cred.PublicKey!, cred.SignCount, isOwner, cred.UserHandle);

        _store[key] = cred with { SignCount = assertResult.Counter };
        return Ok(new { ok = true });
    }

}
