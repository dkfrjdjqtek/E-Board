using WebApplication1.Models;
using System;

// 2025.09.25 Added: 초대 메일 발송 이력 테이블 엔티티 신규 추가 PK만 정의 FK 미생성
namespace WebApplication1.Models
{
    public class InviteAudit
    {
        // PK
        public int Id { get; set; }

        // 대상 사용자 식별자 문자열 길이 제한은 마이그레이션에서 처리
        public string UserId { get; set; } = string.Empty;

        // 발송 대상 이메일
        public string Email { get; set; } = string.Empty;

        // 발송 시각 UTC
        public DateTime SentAtUtc { get; set; }

        // 결과 코드 예 성공 실패
        public string Status { get; set; } = "성공";

        // 오류 메시지 실패 시 저장
        public string? ErrorMessage { get; set; }

        // 상관 아이디 동일 사용자 다중 발송 추적 등 사용
        public string? CorrelationId { get; set; }
    }
}
