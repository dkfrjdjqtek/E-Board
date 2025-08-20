using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Identity;
using WebApplication1.Data;    // ApplicationDbContext
using WebApplication1.Models;  // WebAuthnCredential

namespace WebApplication1.Services
{
    public class WebAuthnService
    {
        private readonly ApplicationDbContext _db;

        public WebAuthnService(ApplicationDbContext db)
        {
            _db = db;
        }

        private static byte[] Sha256(byte[] data)
        {
            using var sha = SHA256.Create();
            return sha.ComputeHash(data);
        }

        /// <summary>
        /// WebAuthn 등록(Attestation) 저장
        /// </summary>
        public async Task<WebAuthnCredential> RegisterCredentialAsync(
            IdentityUser user,
            byte[] rawIdBytes,
            byte[] publicKeyCoseBytes,
            Guid? aaGuid,
            int signCount,
            bool isDiscoverable,
            bool backupEligible,
            bool backedUp,
            IEnumerable<string>? transports,
            string? deviceNickname)
        {
            var entity = new WebAuthnCredential
            {
                Id = Guid.NewGuid(),
                UserId = user.Id,
                CredentialId = rawIdBytes,
                CredentialIdHash = Sha256(rawIdBytes),
                PublicKey = publicKeyCoseBytes,
                AaGuid = aaGuid,
                SignCount = signCount,
                IsDiscoverable = isDiscoverable,
                IsBackupEligible = backupEligible,
                IsBackedUp = backedUp,
                Transports = transports == null ? null : string.Join(",", transports),
                Nickname = deviceNickname,
                CreatedAtUtc = DateTime.UtcNow
            };

            _db.WebAuthnCredentials.Add(entity);
            await _db.SaveChangesAsync();
            return entity;
        }

        /// <summary>
        /// 인증(Assertion) 성공 후 카운터/최근사용 업데이트
        /// </summary>
        public async Task UpdateCounterAsync(Guid credentialPkId, int newSignCount)
        {
            var cred = await _db.WebAuthnCredentials.FindAsync(credentialPkId);
            if (cred == null) return;

            cred.SignCount = newSignCount;
            cred.LastUsedAtUtc = DateTime.UtcNow;
            await _db.SaveChangesAsync();
        }

        /// <summary>
        /// 원본 rawId로 자격증명 검색(해시 비교)
        /// </summary>
        public WebAuthnCredential? FindByRawId(byte[] rawId)
        {
            var hash = Sha256(rawId);
            return _db.WebAuthnCredentials.FirstOrDefault(x => x.CredentialIdHash == hash);
        }
    }
}
