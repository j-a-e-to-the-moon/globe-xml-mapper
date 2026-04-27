using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class MappingOrchestrator
    {
        // 섹션키 → 매퍼 생성 팩토리
        private static readonly Dictionary<string, Func<MappingBase>> MapperFactory = new()
        {
            { "1.1~1.2", () => new Mapping_1_1_1_2() },
            { "1.3.1", () => new Mapping_1_3_1() },
            { "1.3.2.1", () => new Mapping_1_3_2_1() },
            { "1.3.2.2", () => new Mapping_1_3_2_2() },
            { "1.3.3", () => new Mapping_1_3_3() },
            { "1.4", () => new Mapping_1_4() },
            { "2", () => new Mapping_2() },
            { "UTPR", () => new Mapping_Utpr() },
            { "JurCal", () => new Mapping_JurCal() },
            { "EntityCe", () => new Mapping_EntityCe() },
        };

        /// <summary>
        /// 단일 Workbook + _META 숨김시트 기반 매핑.
        /// ControlPanelForm에서 호출.
        /// </summary>
        public List<string> MapWorkbook(string filePath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            using var workbook = new XLWorkbook(filePath);

            // 수식 캐시가 비어있을 수 있으므로(Excel 미저장 상태) 강제 재계산.
            // 실패해도 cached value/원시 입력으로 폴백 가능하므로 throw하지 않음.
            try { workbook.RecalculateAllFormulas(); }
            catch (Exception ex)
            {
                errors.Add($"[경고] 수식 재계산 실패: {ex.Message}. 일부 수식 결과가 누락될 수 있으니 Excel에서 한번 저장 후 재시도하세요.");
            }

            foreach (var (section, sheetName) in TemplateMeta.SheetMap)
            {
                if (!MapperFactory.TryGetValue(section, out var createMapper))
                    continue; // 매퍼 없는 섹션은 스킵 (1.3.3 등 XML 미포함)

                if (!workbook.TryGetWorksheet(sheetName, out var ws))
                    continue; // 시트가 없으면 해당 섹션은 건너뜀

                var mapper = createMapper();
                mapper.Map(ws, globe, errors, sheetName);
            }

            FillMessageSpec(globe);
            EnsureXsdRequiredFields(globe, errors);
            return errors;
        }

        /// <summary>
        /// 매핑 후처리: XSD [R] 필드가 부모 객체와 함께 반드시 emit되도록 보장.
        /// 사용자 입력이 없으면 "0" 기본값으로 채우고 [필수누락] 경고 추가.
        ///
        /// 처리 대상:
        ///   - CEComputation.AdjustedCoveredTax → DeferTaxAdjustAmt [R]
        ///   - OverallComputation [R] children: FANIL, AdjustedFANIL, NetGlobeIncome.Total, IncomeTaxExpense
        ///     (ETRRate, TopUpTaxPercentage는 decimal 0이 항상 emit되도록 Globe.Partials에서 ShouldSerialize 제거됨)
        ///   - SubstanceExclusion이 emit되면 PayrollCost/PayrollMarkUp/TangibleAssetValue/TangibleAssetMarkup 모두 [R]
        ///     (decimal MarkUp은 0 default emit, integer 필드는 빈 값 → "0" 보장)
        /// </summary>
        private static void EnsureXsdRequiredFields(Globe.GlobeOecd globe, List<string> errors)
        {
            if (globe?.GlobeBody?.JurisdictionSection == null) return;

            foreach (var js in globe.GlobeBody.JurisdictionSection)
            {
                var jur = js.Jurisdiction.ToString().ToUpper();

                foreach (var etr in js.GLoBeTax?.Etr ?? Enumerable.Empty<Globe.EtrType>())
                {
                    var comp = etr.EtrStatus?.EtrComputation;
                    if (comp == null) continue;

                    EnsureCeComputationRequired(comp, jur, errors);
                    EnsureOverallComputationRequired(comp.OverallComputation, jur, errors);
                }
            }
        }

        // CEComputation.AdjustedCoveredTax → DeferTaxAdjustAmt [R] 보장
        private static void EnsureCeComputationRequired(
            Globe.EtrComputationType comp, string jur, List<string> errors)
        {
            foreach (var ce in comp.CeComputation ?? new System.Collections.ObjectModel.Collection<Globe.EtrComputationTypeCeComputation>())
            {
                var ceTin = ce.Tin?.Value ?? "?";
                if (ce.AdjustedCoveredTax != null && ce.AdjustedCoveredTax.DeferTaxAdjustAmt == null)
                {
                    ce.AdjustedCoveredTax.DeferTaxAdjustAmt =
                        new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmt
                        {
                            Total = "0",
                            DeferTaxExpense = "0",
                        };
                    errors.Add(
                        $"[필수누락] [{jur}/CE TIN={ceTin}] AdjustedCoveredTax.DeferTaxAdjustAmt 미입력 → 0 기본값으로 보충. " +
                        $"구성기업 계산 3.2.4.2(c) 이연법인세 입력 필요"
                    );
                }
            }
        }

        // OverallComputation [R] 자식 필드 보장
        private static void EnsureOverallComputationRequired(
            Globe.EtrComputationTypeOverallComputation overall, string jur, List<string> errors)
        {
            if (overall == null) return;

            if (string.IsNullOrEmpty(overall.Fanil))
            {
                overall.Fanil = "0";
                errors.Add($"[필수누락] [{jur}] OverallComputation.FANIL 미입력 → 0 기본값. 국가별 계산 시트 B27 (회계상 순이익) 입력 필요");
            }
            if (string.IsNullOrEmpty(overall.AdjustedFanil))
            {
                overall.AdjustedFanil = "0";
                errors.Add($"[필수누락] [{jur}] OverallComputation.AdjustedFANIL 미입력 → 0 기본값. 국가별 계산 시트 O30 (배분 후 회계상 순손익 합계) 입력 필요");
            }
            if (overall.NetGlobeIncome == null)
            {
                overall.NetGlobeIncome = new Globe.EtrComputationTypeOverallComputationNetGlobeIncome { Total = "0" };
                errors.Add($"[필수누락] [{jur}] OverallComputation.NetGlobeIncome 미입력 → 0 기본값");
            }
            else if (string.IsNullOrEmpty(overall.NetGlobeIncome.Total))
            {
                overall.NetGlobeIncome.Total = "0";
                errors.Add($"[필수누락] [{jur}] OverallComputation.NetGlobeIncome.Total 미입력 → 0 기본값");
            }
            if (string.IsNullOrEmpty(overall.IncomeTaxExpense))
            {
                overall.IncomeTaxExpense = "0";
                errors.Add($"[필수누락] [{jur}] OverallComputation.IncomeTaxExpense 미입력 → 0 기본값. 국가별 계산 시트 I27 (법인세비용) 입력 필요");
            }
            // ETRRate, TopUpTaxPercentage는 decimal로 default 0 자동 emit (ShouldSerialize 제거됨)
            // 사용자가 안 채워도 0%로 emit → XSD 통과. 의미 검증은 ValidationUtil이 별도 처리.

            // ExcessProfits [R]
            if (string.IsNullOrEmpty(overall.ExcessProfits))
            {
                overall.ExcessProfits = "0";
                errors.Add($"[필수누락] [{jur}] OverallComputation.ExcessProfits 미입력 → 0 기본값");
            }
            // TopUpTax [R]
            if (string.IsNullOrEmpty(overall.TopUpTax))
            {
                overall.TopUpTax = "0";
                errors.Add($"[필수누락] [{jur}] OverallComputation.TopUpTax 미입력 → 0 기본값");
            }

            // SubstanceExclusion 안의 [R] integer 필드들
            var se = overall.SubstanceExclusion;
            if (se != null)
            {
                if (string.IsNullOrEmpty(se.Total)) se.Total = "0";
                if (string.IsNullOrEmpty(se.PayrollCost)) se.PayrollCost = "0";
                if (string.IsNullOrEmpty(se.TangibleAssetValue)) se.TangibleAssetValue = "0";
                // MarkUp은 decimal — Globe.Partials에서 ShouldSerialize 제거 후 0이라도 emit
            }

            // OverallComputation.AdjustedCoveredTax.DeferTaxAdjustAmt가 emit되면 [R] 모두 보장
            // XSD sequence: Total, DefTaxAmt, DiffCarryValue, GLoBEValue, BefRecastAdjust, TotalAdjust, PreRecast
            // (a) 요약표에서 일부만 채워지면 위치 기반 sequence 위반 발생 → 모두 0 보장
            var dt = overall.AdjustedCoveredTax?.DeferTaxAdjustAmt;
            if (dt != null)
            {
                if (string.IsNullOrEmpty(dt.Total)) dt.Total = "0";
                if (string.IsNullOrEmpty(dt.DefTaxAmt)) dt.DefTaxAmt = "0";
                if (string.IsNullOrEmpty(dt.DiffCarryValue)) dt.DiffCarryValue = "0";
                if (string.IsNullOrEmpty(dt.GLoBeValue)) dt.GLoBeValue = "0";
                if (string.IsNullOrEmpty(dt.BefRecastAdjust)) dt.BefRecastAdjust = "0";
                if (string.IsNullOrEmpty(dt.TotalAdjust)) dt.TotalAdjust = "0";
                if (string.IsNullOrEmpty(dt.PreRecast)) dt.PreRecast = "0";
            }
        }

        #region MessageSpec / DocSpec

        private void FillMessageSpec(Globe.GlobeOecd globe)
        {
            var spec = globe.MessageSpec;
            var fi = globe.GlobeBody?.FilingInfo;

            if (fi?.FilingCe != null)
            {
                spec.TransmittingCountry = fi.FilingCe.ResCountryCode;
                if (!string.IsNullOrWhiteSpace(fi.FilingCe.Tin?.Value))
                    spec.SendingEntityIn = fi.FilingCe.Tin.Value;
            }

            spec.ReceivingCountry = spec.TransmittingCountry;
            spec.MessageType = Globe.MessageTypeEnumType.Gir;

            if (fi?.Period != null && fi.Period.End != default)
                spec.ReportingPeriod = fi.Period.End;

            spec.Timestamp = DateTime.Now;

            if (string.IsNullOrEmpty(spec.MessageRefId))
            {
                var sendCC = spec.TransmittingCountry.ToString().ToUpper();
                var recvCC = spec.ReceivingCountry.ToString().ToUpper();
                var uid = spec.Timestamp.ToString("yyyyMMddHHmmss");
                spec.MessageRefId = $"{sendCC}{spec.ReportingPeriod:yyyy}{recvCC}{uid}";
            }

            FillDocSpecs(globe);
        }

        private void FillDocSpecs(Globe.GlobeOecd globe)
        {
            var sendCC = globe.MessageSpec.TransmittingCountry.ToString().ToUpper();
            var year = globe.MessageSpec.ReportingPeriod.ToString("yyyy");
            var ts = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            if (globe.GlobeBody.FilingInfo != null)
            {
                globe.GlobeBody.FilingInfo.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}FI{ts}",
                };
            }

            if (globe.GlobeBody.GeneralSection != null)
            {
                globe.GlobeBody.GeneralSection.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}GS{ts}",
                };
            }

            int utprIdx = 0;
            foreach (var ua in globe.GlobeBody.UtprAttribution)
            {
                ua.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}UA{utprIdx++}{ts}",
                };
            }

            int jsIdx = 0;
            foreach (var js in globe.GlobeBody.JurisdictionSection)
            {
                js.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}JS{jsIdx++}{ts}",
                };
            }

            int sumIdx = 0;
            foreach (var sm in globe.GlobeBody.Summary)
            {
                sm.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}SM{sumIdx++}{ts}",
                };
            }
        }

        #endregion
    }
}
