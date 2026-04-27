// Globe.cs(XSD 자동생성)는 값형 필드에 Specified 패턴이 없어서
// decimal 기본값 0이 항상 serialize됨. ShouldSerialize* 메서드로 억제.
//
// XmlSerializer 규약: public bool ShouldSerializeXxx() → false 반환 시 Xxx 요소 생략.
//
// ⚠️ XSD [R] 필수 필드는 ShouldSerialize 정의하지 않음 (0이라도 emit 필요).
//   - OverallComputation.TopUpTaxPercentage [R]
//   - SubstanceExclusion.PayrollMarkUp / TangibleAssetMarkup [R] (parent emit 시)
//   → 이 필드들은 emit 강제 (XSD 검증 통과 보장)

namespace Globe
{
    public partial class EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtTransition
    {
        // Year: DateTime default(0001-01-01)이면 사용자 입력 없음 → 직렬화 생략
        // XSD상 required지만 실무상 미기재 허용 (Transition 자체가 optional이라 부모 안 만들면 됨)
        public bool ShouldSerializeYear() => Year != default;
    }

    public partial class EtrTypeElectionArt321C
    {
        // 하위 선택 필드들 — null/빈 문자열이면 <태그 /> 출력 방지
        public bool ShouldSerializeKEquityInvestmentInclusionElection() =>
            !string.IsNullOrEmpty(KEquityInvestmentInclusionElection);

        public bool ShouldSerializeQualOwnerIntentBalance() =>
            !string.IsNullOrEmpty(QualOwnerIntentBalance);

        public bool ShouldSerializeAdditions() => !string.IsNullOrEmpty(Additions);

        public bool ShouldSerializeReductions() => !string.IsNullOrEmpty(Reductions);

        public bool ShouldSerializeOutstandingBalance() =>
            !string.IsNullOrEmpty(OutstandingBalance);
    }
}
