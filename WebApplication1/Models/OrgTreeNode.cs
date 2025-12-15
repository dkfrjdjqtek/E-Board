// 2025.12.12 Added: 조직 트리 멀티 선택 콤보박스용 OrgTreeNode 모델 신규 추가 기타 파일 변경 없음
using System.Collections.Generic;

namespace WebApplication1.Models
{
    /// <summary>
    /// 지사 Branch 부서 Dept 부서원 User 트리 공통으로 쓰는 노드 모델
    /// </summary>
    public class OrgTreeNode
    {
        /// <summary>
        /// 노드 고유 ID (지사코드 부서코드 사번 등)
        /// </summary>
        public string NodeId { get; set; } = string.Empty;

        /// <summary>
        /// 화면에 표시할 이름
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 노드 유형 Branch Dept User 등
        /// </summary>
        public string NodeType { get; set; } = string.Empty;

        /// <summary>
        /// 상위 노드 ID 없으면 null
        /// </summary>
        public string? ParentId { get; set; }

        /// <summary>
        /// 자식 노드 리스트
        /// </summary>
        public List<OrgTreeNode> Children { get; set; } = new List<OrgTreeNode>();
    }
}
