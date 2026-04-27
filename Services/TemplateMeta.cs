namespace GlobeMapper.Services
{
    /// <summary>
    /// main_template_newest.xlsx 메타 정보 (섹션→시트 처리 순서).
    /// </summary>
    public static class TemplateMeta
    {
        // 처리 순서 중요: JurCal이 EntityCe보다 먼저 — JurisdictionSection 선행 생성 보장
        internal static readonly (string section, string sheetName)[] SheetMap = new[]
        {
            ("1.1~1.2", "MNE그룹 정보"),
            ("1.3.1", "최종모기업"),
            ("1.3.2.1", "그룹구조"),
            ("1.3.2.2", "제외기업"),
            ("1.3.3", "그룹구조 변동"),
            ("1.4", "요약"),
            ("2", "적용면제"),
            ("JurCal", "국가별 계산"),
            ("EntityCe", "구성기업 계산"),
            ("UTPR", "UTPR 배분"),
        };
    }
}
