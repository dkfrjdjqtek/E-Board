// Models/WebAuthnCredential.cs
using System;

namespace WebApplication1.Models   // ← 프로젝트 네임스페이스에 맞게
{
    public class WebAuthnCredential
    {
        public Guid Id { get; set; }
        public string UserId { get; set; } = default!;

        public byte[] CredentialId { get; set; } = default!;
        public byte[] CredentialIdHash { get; set; } = default!; // SHA-256(CredentialId)

        public byte[] PublicKey { get; set; } = default!;
        public string CredType { get; set; } = "public-key";
        public Guid? AaGuid { get; set; }
        public byte[]? UserHandle { get; set; }
        public int SignCount { get; set; } = 0;
        public bool IsDiscoverable { get; set; }
        public bool IsBackupEligible { get; set; }
        public bool IsBackedUp { get; set; }
        public string? Transports { get; set; }
        public string? Nickname { get; set; }
        public DateTime CreatedAtUtc { get; set; }
        public DateTime? LastUsedAtUtc { get; set; }
    }
}
