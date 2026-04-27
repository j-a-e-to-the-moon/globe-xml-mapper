using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 시트 3.1~3.2.3.2: 국가별 글로벌최저한세 계산.
    /// 모든 행 위치를 B열 헤더 탐색 + 상대 오프셋으로 처리 (행추가 대응).
    ///
    /// 헤더 기준 오프셋:
    ///   "3.1 국가별"     +1=국가명(O), +2=SubGroup유형(O), +3=SubGroup.TIN(O)
    ///   "3.2.1 실효세율" +2=회계상순손익 행(B=FANIL, I=IncomeTaxExpense, O=ETRRate)
    ///                    data+3=AdjustedFANIL(O), data+5~+30=NetGlobeIncome.Adjustments(26개),
    ///                    data+31=NetGlobeIncome.Total(O)
    ///   "3.2.1.2"        O=AggregrateCurrentTax, +2~+21=AdjustedCoveredTax.Adjustments(20개),
    ///                    +22=AdjustedCoveredTax.Total(O)
    ///   "(b) 이월조정"   +1=PriorYearBalance, +2=GeneratedInRFY,
    ///                    +3=UtilizedInRFY, +4=Remaining (모두 O열)
    ///   "(c) 통합형피지배외국법인"
    ///                    +2행~ B=Jurisdiction, E=SubGroupTIN, J=AggAllocTax (B비면 종료)
    ///   "(a) 요약표"     +1=DefTaxAmt, +2=DiffCarryValue, +3=GLoBEValue, +4=BefRecastAdjust,
    ///                    +5=TotalAdjust, +6=PreRecast, +7=Recast.Lower, +8=Recast.Higher,
    ///                    +9=Total (모두 O열)
    ///   "(b) 조정내역"   +1행~ O=Amount, GIR2501~GIR2516 순서 (O비면 스킵, 16개)
    /// </summary>
    public class Mapping_JurCal : MappingBase
    {
        // AdjustmentItemEnumType: GIR2001~GIR2026 (26개, NetGlobeIncome)
        private static readonly Globe.AdjustmentItemEnumType[] AdjustmentItems =
        {
            Globe.AdjustmentItemEnumType.Gir2001,
            Globe.AdjustmentItemEnumType.Gir2002,
            Globe.AdjustmentItemEnumType.Gir2003,
            Globe.AdjustmentItemEnumType.Gir2004,
            Globe.AdjustmentItemEnumType.Gir2005,
            Globe.AdjustmentItemEnumType.Gir2006,
            Globe.AdjustmentItemEnumType.Gir2007,
            Globe.AdjustmentItemEnumType.Gir2008,
            Globe.AdjustmentItemEnumType.Gir2009,
            Globe.AdjustmentItemEnumType.Gir2010,
            Globe.AdjustmentItemEnumType.Gir2011,
            Globe.AdjustmentItemEnumType.Gir2012,
            Globe.AdjustmentItemEnumType.Gir2013,
            Globe.AdjustmentItemEnumType.Gir2014,
            Globe.AdjustmentItemEnumType.Gir2015,
            Globe.AdjustmentItemEnumType.Gir2016,
            Globe.AdjustmentItemEnumType.Gir2017,
            Globe.AdjustmentItemEnumType.Gir2018,
            Globe.AdjustmentItemEnumType.Gir2019,
            Globe.AdjustmentItemEnumType.Gir2020,
            Globe.AdjustmentItemEnumType.Gir2021,
            Globe.AdjustmentItemEnumType.Gir2022,
            Globe.AdjustmentItemEnumType.Gir2023,
            Globe.AdjustmentItemEnumType.Gir2024,
            Globe.AdjustmentItemEnumType.Gir2025,
            Globe.AdjustmentItemEnumType.Gir2026,
        };

        // DeferredAdjustedTaxEnumType: GIR2501~GIR2516 (16개, DeferTaxAdjustAmt.Adjustments)
        private static readonly Globe.DeferredAdjustedTaxEnumType[] DeferTaxAdjItems =
        {
            Globe.DeferredAdjustedTaxEnumType.Gir2501,
            Globe.DeferredAdjustedTaxEnumType.Gir2502,
            Globe.DeferredAdjustedTaxEnumType.Gir2503,
            Globe.DeferredAdjustedTaxEnumType.Gir2504,
            Globe.DeferredAdjustedTaxEnumType.Gir2505,
            Globe.DeferredAdjustedTaxEnumType.Gir2506,
            Globe.DeferredAdjustedTaxEnumType.Gir2507,
            Globe.DeferredAdjustedTaxEnumType.Gir2508,
            Globe.DeferredAdjustedTaxEnumType.Gir2509,
            Globe.DeferredAdjustedTaxEnumType.Gir2510,
            Globe.DeferredAdjustedTaxEnumType.Gir2511,
            Globe.DeferredAdjustedTaxEnumType.Gir2512,
            Globe.DeferredAdjustedTaxEnumType.Gir2513,
            Globe.DeferredAdjustedTaxEnumType.Gir2514,
            Globe.DeferredAdjustedTaxEnumType.Gir2515,
            Globe.DeferredAdjustedTaxEnumType.Gir2516,
        };

        // FinalAdjustedTaxEnumType: GIR2701~GIR2720 (20개, AdjustedCoveredTax)
        private static readonly Globe.FinalAdjustedTaxEnumType[] CoveredTaxAdjItems =
        {
            Globe.FinalAdjustedTaxEnumType.Gir2701,
            Globe.FinalAdjustedTaxEnumType.Gir2702,
            Globe.FinalAdjustedTaxEnumType.Gir2703,
            Globe.FinalAdjustedTaxEnumType.Gir2704,
            Globe.FinalAdjustedTaxEnumType.Gir2705,
            Globe.FinalAdjustedTaxEnumType.Gir2706,
            Globe.FinalAdjustedTaxEnumType.Gir2707,
            Globe.FinalAdjustedTaxEnumType.Gir2708,
            Globe.FinalAdjustedTaxEnumType.Gir2709,
            Globe.FinalAdjustedTaxEnumType.Gir2710,
            Globe.FinalAdjustedTaxEnumType.Gir2711,
            Globe.FinalAdjustedTaxEnumType.Gir2712,
            Globe.FinalAdjustedTaxEnumType.Gir2713,
            Globe.FinalAdjustedTaxEnumType.Gir2714,
            Globe.FinalAdjustedTaxEnumType.Gir2715,
            Globe.FinalAdjustedTaxEnumType.Gir2716,
            Globe.FinalAdjustedTaxEnumType.Gir2717,
            Globe.FinalAdjustedTaxEnumType.Gir2718,
            Globe.FinalAdjustedTaxEnumType.Gir2719,
            Globe.FinalAdjustedTaxEnumType.Gir2720,
        };

        public Mapping_JurCal()
            : base(null) { }

        // 현재 처리 중인 블록 범위 (FindRow 스코프)
        private int _blockStart = 1;
        private int _blockEnd = -1;
        // ReportingPeriod 종료일 (라벨→연도 매핑용, "신고대상 사업연도" 등)
        private DateTime _filingPeriodEnd = default;

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            // 라벨→연도 매핑 컨텍스트 (Map323DeemedDistTax 등에서 사용)
            _filingPeriodEnd = globe.GlobeBody?.FilingInfo?.Period?.End ?? default;

            // 세로 스택된 "3.1 국가별" 블록을 모두 찾아 각각 처리
            var blockStarts = FindAllBlockStarts(ws);
            if (blockStarts.Count == 0)
                return;

            var lastUsedRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            for (int i = 0; i < blockStarts.Count; i++)
            {
                _blockStart = blockStarts[i];
                _blockEnd = (i + 1 < blockStarts.Count) ? blockStarts[i + 1] - 1 : lastUsedRow;
                MapJurisdiction(ws, globe, errors, fileName);
            }

            // 블록 컨텍스트 리셋
            _blockStart = 1;
            _blockEnd = -1;
            _filingPeriodEnd = default;
        }

        /// <summary>
        /// 전체 시트에서 "3.1 국가별 글로벌" 헤더 행 모두 찾기 (각각 블록 시작점).
        /// 주의: "3.1 국가별"만으로 matching하면 "3.2.3.1 국가별 선택..."도 잡혀 블록이 과분할됨.
        /// </summary>
        private static List<int> FindAllBlockStarts(IXLWorksheet ws)
        {
            var result = new List<int>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            for (int r = 1; r <= lastRow; r++)
            {
                var txt = ws.Cell(r, 2).GetString()?.Trim() ?? "";
                if (txt.Contains("3.1 국가별 글로벌"))
                    result.Add(r);
            }
            return result;
        }

        // 템플릿 라벨("신고대상 사업연도", "직전연도", "직전 N년차")을 연도로 변환.
        // 기준은 _filingPeriodEnd.Year (FilingInfo.Period.End). 라벨 미인식 시 false.
        private bool TryMapLabelToYear(string label, out DateTime year)
        {
            year = default;
            if (_filingPeriodEnd == default || string.IsNullOrEmpty(label))
                return false;
            var rfy = _filingPeriodEnd.Year;
            int? offset = null;
            if (label.Contains("신고대상")) offset = 0;
            else if (label.Contains("직전연도") || label == "직전 1년차") offset = -1;
            else if (label.Contains("직전 2년차") || label.Contains("직전 전")) offset = -2;
            else if (label.Contains("직전 3년차")) offset = -3;
            else if (label.Contains("직전 4년차")) offset = -4;
            else if (label.Contains("직전 5년차")) offset = -5;
            if (!offset.HasValue) return false;
            year = new DateTime(rfy + offset.Value, 1, 1);
            return true;
        }

        // B열에서 contains 텍스트를 포함하는 행 반환 (-1 = 없음).
        // fromRow/toRow 생략 시 현재 블록 범위(_blockStart, _blockEnd) 사용.
        private int FindRow(IXLWorksheet ws, string contains, int fromRow = 0, int toRow = 0)
        {
            if (fromRow <= 0)
                fromRow = _blockStart;
            int end =
                toRow > 0
                    ? toRow
                    : (_blockEnd > 0 ? _blockEnd : ws.LastRowUsed()?.RowNumber() ?? 300);
            for (int r = fromRow; r <= end; r++)
            {
                var txt = ws.Cell(r, 2).GetString()?.Trim() ?? "";
                if (txt.Contains(contains))
                    return r;
            }
            return -1;
        }

        // ─── 3.1 국가명 + JurisdictionSection 생성 ───────────────────────
        private void MapJurisdiction(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            var row31 = FindRow(ws, "3.1 국가별");
            if (row31 < 0)
                return;

            var jurRaw = ws.Cell(row31 + 1, 15).GetString()?.Trim(); // O: +1 국가명 (O5)
            if (string.IsNullOrEmpty(jurRaw))
                return;

            if (!TryParseEnum<Globe.CountryCodeType>(jurRaw, out var countryCode))
            {
                errors.Add($"[{fileName}] [3.1] 국가코드 '{jurRaw}' 파싱 실패");
                return;
            }

            // JurisdictionSection 찾기 또는 생성
            var js = globe.GlobeBody.JurisdictionSection.FirstOrDefault(s =>
                s.Jurisdiction == countryCode
            );
            if (js == null)
            {
                js = new Globe.GlobeBodyTypeJurisdictionSection
                {
                    Jurisdiction = countryCode,
                    GLoBeTax = new Globe.GlobeTax(),
                };
                js.RecJurCode.Add(countryCode);
                globe.GlobeBody.JurisdictionSection.Add(js);
            }

            // ETR 찾기 또는 생성
            var etr = js.GLoBeTax.Etr.FirstOrDefault();
            if (etr == null)
            {
                etr = new Globe.EtrType { EtrStatus = new Globe.EtrTypeEtrStatus() };
                js.GLoBeTax.Etr.Add(etr);
            }

            // O: +2 하위그룹유형, +3 TIN
            var subGroupTypeRaw = ws.Cell(row31 + 2, 15).GetString()?.Trim(); // O6
            var subGroupTinRaw = ws.Cell(row31 + 3, 15).GetString()?.Trim(); // O7
            if (!string.IsNullOrEmpty(subGroupTypeRaw) || !string.IsNullOrEmpty(subGroupTinRaw))
            {
                var subGroup = new Globe.EtrTypeSubGroup();
                subGroup.Tin = ParseTin(
                    string.IsNullOrEmpty(subGroupTinRaw) ? "NOTIN" : subGroupTinRaw
                );
                if (!string.IsNullOrEmpty(subGroupTypeRaw))
                {
                    foreach (
                        var code in subGroupTypeRaw.Split(
                            ',',
                            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
                        )
                    )
                        SetEnum<Globe.EtrTypeofSubGroupEnumType>(
                            code,
                            v => subGroup.TypeofSubGroup.Add(v),
                            errors,
                            fileName,
                            new MappingEntry { Cell = "3.1+2/O", Label = "[3.1] 하위그룹유형" }
                        );
                }
                etr.SubGroup = subGroup;
            }

            Map321OverallComputation(ws, etr, errors, fileName);
            Map322DeferTaxAdjust(ws, etr, errors, fileName);
            Map322Transition(ws, etr, errors, fileName);
            Map322LossCarryback(ws, etr, errors, fileName);
            Map322Recapture(ws, etr, errors, fileName);
            Map323Election(ws, etr, errors, fileName);
            Map323DeemedDistTax(ws, etr, errors, fileName);

            // 3.2.4.4 국제해운소득·결손 제외 — 국가별 계산 시트 내부에 통합됨
            MapShippingIncome(ws, etr, errors, fileName);

            // 3.3 추가세액 계산 — 국가별 계산 시트 내부에 통합됨
            Map33TopUpTaxCalc(ws, etr, errors, fileName);
        }

        // ─── 3.2.1 OverallComputation ────────────────────────────────────
        private void Map321OverallComputation(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            // ── 3.2.1 실효세율 헤더 → +2 = 회계상순손익 행 ─────────────────
            var row321 = FindRow(ws, "3.2.1 실효세율");
            if (row321 < 0)
                return;
            var dRow = row321 + 2; // 회계상순손익(FANIL) 행

            var fanilRaw = ws.Cell(dRow, 2).GetString()?.Trim(); // B: FANIL
            var incTaxRaw = ws.Cell(dRow, 9).GetString()?.Trim(); // I: IncomeTaxExpense
            var etrRateRaw = ws.Cell(dRow, 15).GetString()?.Trim(); // O: ETRRate
            var adjFanilRaw = ws.Cell(dRow + 3, 15).GetString()?.Trim(); // O+3: AdjustedFANIL
            var netGlobeRaw = ws.Cell(dRow + 31, 15).GetString()?.Trim(); // O+31: NetGlobeIncome.Total

            // ── 3.2.1.2 조정대상조세 헤더 ────────────────────────────────────
            // +0: 3.2.1.2 헤더, +1: (a)소계헤더, +2: AggregrateCurrentTax,
            // +3: "2.조정사항" 서브헤더, +4~+23: (a)~(t) 20개, +24: Total
            var row3212 = FindRow(ws, "3.2.1.2");
            var aggTaxRaw = row3212 >= 0 ? ws.Cell(row3212 + 2, 15).GetString()?.Trim() : null; // O+2: AggregrateCurrentTax (O62)
            var adjTaxRaw = row3212 >= 0 ? ws.Cell(row3212 + 24, 15).GetString()?.Trim() : null; // O+24: Total (O84)

            // ── (b) 이월조정대상조세 헤더 ─────────────────────────────────────
            var rowExcess = FindRow(ws, "(b) 이월조정대상조세");
            var exPriorRaw = rowExcess >= 0 ? ws.Cell(rowExcess + 1, 15).GetString()?.Trim() : null; // O87
            var exGenRaw = rowExcess >= 0 ? ws.Cell(rowExcess + 2, 15).GetString()?.Trim() : null; // O88
            var exUtilRaw = rowExcess >= 0 ? ws.Cell(rowExcess + 3, 15).GetString()?.Trim() : null; // O89
            var exRemRaw = rowExcess >= 0 ? ws.Cell(rowExcess + 4, 15).GetString()?.Trim() : null; // O90

            bool hasData =
                !string.IsNullOrEmpty(fanilRaw)
                || !string.IsNullOrEmpty(incTaxRaw)
                || !string.IsNullOrEmpty(etrRateRaw)
                || !string.IsNullOrEmpty(adjFanilRaw)
                || !string.IsNullOrEmpty(netGlobeRaw)
                || !string.IsNullOrEmpty(aggTaxRaw)
                || !string.IsNullOrEmpty(adjTaxRaw);
            if (!hasData)
                return;

            if (etr.EtrStatus == null)
                etr.EtrStatus = new Globe.EtrTypeEtrStatus();
            if (etr.EtrStatus.EtrComputation == null)
                etr.EtrStatus.EtrComputation = new Globe.EtrComputationType();

            var overall =
                etr.EtrStatus.EtrComputation.OverallComputation
                ?? new Globe.EtrComputationTypeOverallComputation();
            etr.EtrStatus.EtrComputation.OverallComputation = overall;

            // a. FANIL (B)
            if (!string.IsNullOrEmpty(fanilRaw))
                overall.Fanil = fanilRaw;

            // AdjustedFANIL (O+3)
            if (!string.IsNullOrEmpty(adjFanilRaw))
                overall.AdjustedFanil = adjFanilRaw;

            // NetGlobeIncome.Adjustments (O+5 ~ O+30, 26개)
            overall.NetGlobeIncome ??=
                new Globe.EtrComputationTypeOverallComputationNetGlobeIncome();
            for (int i = 0; i < AdjustmentItems.Length; i++)
            {
                var amount = ws.Cell(dRow + 5 + i, 15).GetString()?.Trim();
                if (!HasNonZeroValue(amount))
                    continue;
                overall.NetGlobeIncome.Adjustments.Add(
                    new Globe.EtrComputationTypeOverallComputationNetGlobeIncomeAdjustments
                    {
                        Amount = amount,
                        AdjustmentItem = AdjustmentItems[i],
                    }
                );
            }

            // NetGlobeIncome.Total (O+31)
            if (!string.IsNullOrEmpty(netGlobeRaw))
                overall.NetGlobeIncome.Total = netGlobeRaw;

            // c. IncomeTaxExpense (I)
            if (!string.IsNullOrEmpty(incTaxRaw))
                overall.IncomeTaxExpense = incTaxRaw;

            // d. AdjustedCoveredTax
            if (row3212 >= 0)
            {
                overall.AdjustedCoveredTax ??=
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTax();

                // AggregrateCurrentTax (O)
                if (!string.IsNullOrEmpty(aggTaxRaw))
                    overall.AdjustedCoveredTax.AggregrateCurrentTax = aggTaxRaw;

                // Adjustments (+2 ~ +21, 20개) → GIR2701~GIR2720
                for (int i = 0; i < CoveredTaxAdjItems.Length; i++)
                {
                    var amount = ws.Cell(row3212 + 4 + i, 15).GetString()?.Trim();
                    if (!HasNonZeroValue(amount))
                        continue;
                    overall.AdjustedCoveredTax.Adjustments.Add(
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxAdjustments
                        {
                            Amount = amount,
                            AdjustmentItem = CoveredTaxAdjItems[i],
                        }
                    );
                }

                // Total (+22)
                if (!string.IsNullOrEmpty(adjTaxRaw))
                    overall.AdjustedCoveredTax.Total = adjTaxRaw;

                // (c) 통합형피지배외국법인: rowCfc+2~, B비면 종료
                // 실제 CfcJur 추가 시점에 TransBlendCfc 생성 (빈 태그 방지)
                var rowCfc = FindRow(ws, "(c) 통합형피지배외국법인");
                if (rowCfc >= 0)
                {
                    for (int row = rowCfc + 2; ; row++)
                    {
                        var cfcJurRaw = ws.Cell(row, 2).GetString()?.Trim();
                        if (string.IsNullOrEmpty(cfcJurRaw))
                            break;
                        if (!TryParseEnum<Globe.CountryCodeType>(cfcJurRaw, out var cfcJur))
                        {
                            errors.Add(
                                $"[{fileName}] [3.2.1(c)] 행{row} 국가코드 '{cfcJurRaw}' 파싱 실패"
                            );
                            continue;
                        }
                        var tinRaw = ws.Cell(row, 5).GetString()?.Trim();
                        var amtRaw = ws.Cell(row, 10).GetString()?.Trim();
                        if (string.IsNullOrEmpty(tinRaw) && string.IsNullOrEmpty(amtRaw))
                            continue;

                        overall.AdjustedCoveredTax.TransBlendCfc ??=
                            new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxTransBlendCfc();

                        overall.AdjustedCoveredTax.TransBlendCfc.CfcJur.Add(
                            new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxTransBlendCfcCfcJur
                            {
                                Jurisdiction = cfcJur,
                                Allocation =
                                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxTransBlendCfcCfcJurAllocation
                                    {
                                        SubGroupTin = string.IsNullOrEmpty(tinRaw)
                                            ? null
                                            : ParseTin(tinRaw),
                                        AggAllocTax = amtRaw,
                                    },
                            }
                        );
                    }

                    // TransBlendCfc.Total [R] = ΣAllocation.AggAllocTax (decimal 합계)
                    if (overall.AdjustedCoveredTax.TransBlendCfc != null)
                    {
                        decimal sum = 0m;
                        foreach (var cj in overall.AdjustedCoveredTax.TransBlendCfc.CfcJur)
                        {
                            if (decimal.TryParse(
                                    cj.Allocation?.AggAllocTax,
                                    System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    out var v))
                                sum += v;
                        }
                        overall.AdjustedCoveredTax.TransBlendCfc.Total =
                            sum.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                }
            }

            // f. ExcessNegTaxExpense ((b) 이월조정대상조세)
            if (
                !string.IsNullOrEmpty(exPriorRaw)
                || !string.IsNullOrEmpty(exGenRaw)
                || !string.IsNullOrEmpty(exUtilRaw)
                || !string.IsNullOrEmpty(exRemRaw)
            )
            {
                overall.ExcessNegTaxExpense ??=
                    new Globe.EtrComputationTypeOverallComputationExcessNegTaxExpense();
                if (!string.IsNullOrEmpty(exPriorRaw))
                    overall.ExcessNegTaxExpense.PriorYearBalance = exPriorRaw;
                if (!string.IsNullOrEmpty(exGenRaw))
                    overall.ExcessNegTaxExpense.GeneratedInRfy = exGenRaw;
                if (!string.IsNullOrEmpty(exUtilRaw))
                    overall.ExcessNegTaxExpense.UtilizedInRfy = exUtilRaw;
                if (!string.IsNullOrEmpty(exRemRaw))
                    overall.ExcessNegTaxExpense.Remaining = exRemRaw;
            }

            // e. ETRRate (O) — 퍼센트(%) 입력 시 /100
            if (!string.IsNullOrEmpty(etrRateRaw))
            {
                var rate = ParsePercentage(etrRateRaw);
                if (rate.HasValue)
                    overall.EtrRate = rate.Value;
                else
                    errors.Add($"[{fileName}] [3.2.1] ETRRate 형식 오류: '{etrRateRaw}' (0~1 사이 decimal 입력, 예: 0.15)");
            }
        }

        // ─── 3.2.2.3 최초적용연도에 대한 특례 → DeferTaxAdjustAmt.Transition[] ──
        private void Map322Transition(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            var rowHeader = FindRow(ws, "3.2.2.3");
            if (rowHeader < 0)
                return;

            // 레이아웃:
            //  +0: "3.2.2.3 최초적용연도에 대한 특례" 헤더
            //  +1: "1. 최초적용연도" 레이블 행 — O열에 연도값
            //  +2: "(a)..." 서브헤더
            //  +3: "이 연 법 인 세 부 채" 그룹헤더
            //  +4: [B="1. DTLStart"] [K="2. DTLRecast"] 컬럼헤더
            //  +5: 데이터행 — B=DTLStart, K=DTLRecast
            //  +6: (빈 행)
            //  +7: "이 연 법 인 세 자 산" 그룹헤더
            //  +8: [B="3."] [F="4."] [K="5."] [O="6."] 컬럼헤더
            //  +9: 데이터행 — B=DTAStart, F=DTARecast, K=DTAExcl, O=DTATotal
            var dtlStartRaw = ws.Cell(rowHeader + 5, 2).GetString()?.Trim(); // 1. B열 (B155)
            var dtlRecastRaw = ws.Cell(rowHeader + 5, 11).GetString()?.Trim(); // 2. K열 (K155)
            var dtaStartRaw = ws.Cell(rowHeader + 9, 2).GetString()?.Trim(); // 3. B열 (B159)
            var dtaRecastRaw = ws.Cell(rowHeader + 9, 6).GetString()?.Trim(); // 4. F열 (F159)
            var dtaExclRaw = ws.Cell(rowHeader + 9, 11).GetString()?.Trim(); // 5. K열 (K159)
            var dtaTotalRaw = ws.Cell(rowHeader + 9, 15).GetString()?.Trim(); // 6. O열 (O159)

            // (b) 처분기업 명세 — "(b)" 서브헤더부터 B열 비면 종료
            var rowDisposal = FindRow(ws, "(b)", rowHeader);
            var disposals =
                new List<Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtTransitionDisposal>();
            if (rowDisposal >= 0)
            {
                // +1: 열 헤더 행 ("1. 처분기업 소재지국" 등) → 스킵, +2부터 데이터
                for (int row = rowDisposal + 2; ; row++)
                {
                    var ccRaw = ws.Cell(row, 2).GetString()?.Trim(); // B: ResCountryCode
                    if (string.IsNullOrEmpty(ccRaw))
                        break;
                    if (!TryParseEnum<Globe.CountryCodeType>(ccRaw, out var cc))
                    {
                        errors.Add(
                            $"[{fileName}] [3.2.2.3(b)] 행{row} 국가코드 '{ccRaw}' 파싱 실패"
                        );
                        continue;
                    }
                    var taxPaidRaw = ws.Cell(row, 4).GetString()?.Trim(); // D
                    var netDtadtlRaw = ws.Cell(row, 8).GetString()?.Trim(); // H
                    var carryValRaw = ws.Cell(row, 11).GetString()?.Trim(); // K
                    var dtadtlRaw = ws.Cell(row, 14).GetString()?.Trim(); // N

                    var d =
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtTransitionDisposal
                        {
                            ResCountryCode = cc,
                        };
                    if (!string.IsNullOrEmpty(taxPaidRaw))
                        d.TaxPaid = taxPaidRaw;
                    if (!string.IsNullOrEmpty(netDtadtlRaw))
                        d.NetDtadtl = netDtadtlRaw;
                    if (!string.IsNullOrEmpty(carryValRaw))
                        d.CarryingValue = carryValRaw;
                    if (!string.IsNullOrEmpty(dtadtlRaw))
                        d.Dtadtl = dtadtlRaw;
                    disposals.Add(d);
                }
            }

            bool hasData =
                !string.IsNullOrEmpty(dtlStartRaw)
                || !string.IsNullOrEmpty(dtlRecastRaw)
                || !string.IsNullOrEmpty(dtaStartRaw)
                || !string.IsNullOrEmpty(dtaRecastRaw)
                || !string.IsNullOrEmpty(dtaExclRaw)
                || !string.IsNullOrEmpty(dtaTotalRaw)
                || disposals.Count > 0;
            if (!hasData)
                return;

            if (etr.EtrStatus == null)
                etr.EtrStatus = new Globe.EtrTypeEtrStatus();
            if (etr.EtrStatus.EtrComputation == null)
                etr.EtrStatus.EtrComputation = new Globe.EtrComputationType();
            var overall =
                etr.EtrStatus.EtrComputation.OverallComputation
                ?? new Globe.EtrComputationTypeOverallComputation();
            etr.EtrStatus.EtrComputation.OverallComputation = overall;
            overall.AdjustedCoveredTax ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTax();
            overall.AdjustedCoveredTax.DeferTaxAdjustAmt ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmt();
            var dt = overall.AdjustedCoveredTax.DeferTaxAdjustAmt;

            // Year: +1행 O열 (최초적용연도 사업연도 개시일)
            var yearRaw = ws.Cell(rowHeader + 1, 15).GetString()?.Trim(); // O151
            DateTime? year = null;
            if (!string.IsNullOrEmpty(yearRaw))
            {
                if (int.TryParse(yearRaw, out var yearNum))
                    year = new DateTime(yearNum, 1, 1);
                else if (DateTime.TryParse(yearRaw, out var yearDate))
                    year = yearDate;
                else
                    errors.Add($"[{fileName}] [3.2.2.3] 최초적용연도 파싱 실패: '{yearRaw}'");
            }

            var transition =
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtTransition();
            if (year.HasValue)
                transition.Year = year.Value;

            if (!string.IsNullOrEmpty(dtlStartRaw))
                transition.DeferredTaxLiabilityStart = dtlStartRaw;
            if (!string.IsNullOrEmpty(dtlRecastRaw))
                transition.DeferredTaxLiabilityRecast = dtlRecastRaw;

            if (
                !string.IsNullOrEmpty(dtaStartRaw)
                || !string.IsNullOrEmpty(dtaRecastRaw)
                || !string.IsNullOrEmpty(dtaExclRaw)
                || !string.IsNullOrEmpty(dtaTotalRaw)
            )
            {
                transition.DeferredTaxAssets ??=
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtTransitionDeferredTaxAssets();
                if (!string.IsNullOrEmpty(dtaStartRaw))
                    transition.DeferredTaxAssets.DeferredTaxAssetStart = dtaStartRaw;
                if (!string.IsNullOrEmpty(dtaRecastRaw))
                    transition.DeferredTaxAssets.DeferredTaxAssetRecast = dtaRecastRaw;
                if (!string.IsNullOrEmpty(dtaExclRaw))
                    transition.DeferredTaxAssets.DeferredTaxAssetExcluded = dtaExclRaw;
                if (!string.IsNullOrEmpty(dtaTotalRaw))
                    transition.DeferredTaxAssets.Total = dtaTotalRaw;
            }

            foreach (var d in disposals)
                transition.Disposal.Add(d);

            dt.Transition.Add(transition);
        }

        // ─── 3.2.2 결손금 소급공제 → PostFilingAdjust ────────────────────
        private void Map322LossCarryback(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            var rowHeader = FindRow(ws, "(c) 결손금 소급공제");
            if (rowHeader < 0)
                return;

            var deferAmts = new List<(DateTime year, string amount)>();
            var refundAmts = new List<(DateTime year, string amount)>();
            string deferTotal = null,
                refundTotal = null;

            // +1 = 열 헤더 행 → +2부터 데이터
            for (int row = rowHeader + 2; ; row++)
            {
                var bVal = ws.Cell(row, 2).GetString()?.Trim();
                if (string.IsNullOrEmpty(bVal))
                    break;

                var fVal = ws.Cell(row, 6).GetString()?.Trim();
                var jVal = ws.Cell(row, 10).GetString()?.Trim();

                if (bVal == "합계")
                {
                    deferTotal = fVal;
                    refundTotal = jVal;
                    break;
                }

                // B열 = 연도 숫자 (예: 2024)
                if (!int.TryParse(bVal, out var yearNum))
                {
                    errors.Add(
                        $"[{fileName}] [(c) 결손금 소급공제] 행{row} 연도 파싱 실패: '{bVal}'"
                    );
                    continue;
                }
                var yearDate = new DateTime(yearNum, 1, 1);

                if (!string.IsNullOrEmpty(fVal))
                    deferAmts.Add((yearDate, fVal));
                if (!string.IsNullOrEmpty(jVal))
                    refundAmts.Add((yearDate, jVal));
            }

            bool hasData =
                deferAmts.Count > 0
                || refundAmts.Count > 0
                || !string.IsNullOrEmpty(deferTotal)
                || !string.IsNullOrEmpty(refundTotal);
            if (!hasData)
                return;

            if (etr.EtrStatus == null)
                etr.EtrStatus = new Globe.EtrTypeEtrStatus();
            if (etr.EtrStatus.EtrComputation == null)
                etr.EtrStatus.EtrComputation = new Globe.EtrComputationType();
            var overall =
                etr.EtrStatus.EtrComputation.OverallComputation
                ?? new Globe.EtrComputationTypeOverallComputation();
            etr.EtrStatus.EtrComputation.OverallComputation = overall;

            overall.AdjustedCoveredTax ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTax();
            overall.AdjustedCoveredTax.PostFilingAdjust ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxPostFilingAdjust();
            var pfa = overall.AdjustedCoveredTax.PostFilingAdjust;

            // DeferTaxAsset
            if (deferAmts.Count > 0 || !string.IsNullOrEmpty(deferTotal))
            {
                pfa.DeferTaxAsset ??=
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxPostFilingAdjustDeferTaxAsset();
                foreach (var (yr, amt) in deferAmts)
                    pfa.DeferTaxAsset.AmountAttributed.Add(
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxPostFilingAdjustDeferTaxAssetAmountAttributed
                        {
                            Year = yr,
                            Amount = amt,
                        }
                    );
                if (!string.IsNullOrEmpty(deferTotal))
                    pfa.DeferTaxAsset.Total = deferTotal;
            }

            // CoveredTaxRefund
            if (refundAmts.Count > 0 || !string.IsNullOrEmpty(refundTotal))
            {
                pfa.CoveredTaxRefund ??=
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxPostFilingAdjustCoveredTaxRefund();
                foreach (var (yr, amt) in refundAmts)
                    pfa.CoveredTaxRefund.AmountAttributed.Add(
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxPostFilingAdjustCoveredTaxRefundAmountAttributed
                        {
                            Year = yr,
                            Amount = amt,
                        }
                    );
                if (!string.IsNullOrEmpty(refundTotal))
                    pfa.CoveredTaxRefund.Total = refundTotal;
            }
        }

        // ─── 3.2.2.2 환입금액 계산 → Adjustments[GIR2509].RecaptureDeferred ─
        private void Map322Recapture(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            // (a) 헤더: O열 +1~+3
            var rowA = FindRow(ws, "(a) 이연법인세부채 환입");
            if (rowA < 0)
                return;

            var dtlMinus5Raw = ws.Cell(rowA + 1, 15).GetString()?.Trim(); // 1. DTLRFYMinus5 (O141)
            var recapMinus5Raw = ws.Cell(rowA + 2, 15).GetString()?.Trim(); // 2. RecapDTLRFYMinus5 (O142)
            var dtlRfyRaw = ws.Cell(rowA + 3, 15).GetString()?.Trim(); // 3. DTLRFY (O143)

            // (b) 헤더: J=신고대상(10), N=직전(14), +2=a, +3=b, +4=c
            var rowB = FindRow(ws, "(b) 총 이연법인세부채 환입", rowA);
            string rfyPreTrans = null,
                rfyOutBal = null,
                rfyUnjust = null;
            string priorPreTrans = null,
                priorOutBal = null,
                priorUnjust = null;
            if (rowB >= 0)
            {
                rfyPreTrans = ws.Cell(rowB + 2, 10).GetString()?.Trim(); // J146
                priorPreTrans = ws.Cell(rowB + 2, 14).GetString()?.Trim(); // N146
                rfyOutBal = ws.Cell(rowB + 3, 10).GetString()?.Trim(); // J147
                priorOutBal = ws.Cell(rowB + 3, 14).GetString()?.Trim(); // N147
                rfyUnjust = ws.Cell(rowB + 4, 10).GetString()?.Trim(); // J148
                priorUnjust = ws.Cell(rowB + 4, 14).GetString()?.Trim(); // N148
            }

            bool hasData =
                !string.IsNullOrEmpty(dtlMinus5Raw)
                || !string.IsNullOrEmpty(recapMinus5Raw)
                || !string.IsNullOrEmpty(dtlRfyRaw)
                || !string.IsNullOrEmpty(rfyPreTrans)
                || !string.IsNullOrEmpty(rfyOutBal)
                || !string.IsNullOrEmpty(rfyUnjust)
                || !string.IsNullOrEmpty(priorPreTrans)
                || !string.IsNullOrEmpty(priorOutBal)
                || !string.IsNullOrEmpty(priorUnjust);
            if (!hasData)
                return;

            if (
                etr.EtrStatus
                    ?.EtrComputation
                    ?.OverallComputation
                    ?.AdjustedCoveredTax
                    ?.DeferTaxAdjustAmt == null
            )
                return; // DeferTaxAdjustAmt가 없으면 건너뜀

            var dt = etr.EtrStatus
                .EtrComputation
                .OverallComputation
                .AdjustedCoveredTax
                .DeferTaxAdjustAmt;

            // 기존 GIR2509 항목 찾기 또는 생성
            var entry = dt.Adjustments.FirstOrDefault(a =>
                a.AdjustmentItem == Globe.DeferredAdjustedTaxEnumType.Gir2509
            );
            if (entry == null)
            {
                entry =
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustments
                    {
                        AdjustmentItem = Globe.DeferredAdjustedTaxEnumType.Gir2509,
                    };
                dt.Adjustments.Add(entry);
            }

            entry.RecaptureDeferred ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentsRecaptureDeferred();
            var rd = entry.RecaptureDeferred;

            if (!string.IsNullOrEmpty(dtlMinus5Raw))
                rd.DtlrfyMinus5 = dtlMinus5Raw;
            if (!string.IsNullOrEmpty(recapMinus5Raw))
                rd.RecapDtlrfyMinus5 = recapMinus5Raw;
            if (!string.IsNullOrEmpty(dtlRfyRaw))
                rd.Dtlrfy = dtlRfyRaw;

            bool hasAggDtl =
                !string.IsNullOrEmpty(rfyPreTrans)
                || !string.IsNullOrEmpty(rfyOutBal)
                || !string.IsNullOrEmpty(rfyUnjust)
                || !string.IsNullOrEmpty(priorPreTrans)
                || !string.IsNullOrEmpty(priorOutBal)
                || !string.IsNullOrEmpty(priorUnjust);
            if (hasAggDtl)
            {
                rd.AggregateDtl ??=
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentsRecaptureDeferredAggregateDtl();

                // ReportingFiscalYear (J열, 신고대상)
                if (
                    !string.IsNullOrEmpty(rfyPreTrans)
                    || !string.IsNullOrEmpty(rfyOutBal)
                    || !string.IsNullOrEmpty(rfyUnjust)
                )
                {
                    rd.AggregateDtl.ReportingFiscalYear ??=
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentsRecaptureDeferredAggregateDtlReportingFiscalYear();
                    if (!string.IsNullOrEmpty(rfyPreTrans))
                        rd.AggregateDtl.ReportingFiscalYear.AmountPreTransition = rfyPreTrans;
                    if (!string.IsNullOrEmpty(rfyOutBal))
                        rd.AggregateDtl.ReportingFiscalYear.AmountOutBalance = rfyOutBal;
                    if (!string.IsNullOrEmpty(rfyUnjust))
                        rd.AggregateDtl.ReportingFiscalYear.AmountUnjustified = rfyUnjust;
                }

                // PriorFiscalYear (N열, 직전)
                if (
                    !string.IsNullOrEmpty(priorPreTrans)
                    || !string.IsNullOrEmpty(priorOutBal)
                    || !string.IsNullOrEmpty(priorUnjust)
                )
                {
                    rd.AggregateDtl.PriorFiscalYear ??=
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentsRecaptureDeferredAggregateDtlPriorFiscalYear();
                    if (!string.IsNullOrEmpty(priorPreTrans))
                        rd.AggregateDtl.PriorFiscalYear.AmountPreTransition = priorPreTrans;
                    if (!string.IsNullOrEmpty(priorOutBal))
                        rd.AggregateDtl.PriorFiscalYear.AmountOutBalance = priorOutBal;
                    if (!string.IsNullOrEmpty(priorUnjust))
                        rd.AggregateDtl.PriorFiscalYear.AmountUnjustified = priorUnjust;
                }
            }
        }

        // ─── 3.2.2.1 DeferTaxAdjustAmt ──────────────────────────────────
        private void Map322DeferTaxAdjust(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            // (a) 요약표 헤더 탐색
            var rowSummary = FindRow(ws, "(a) 요약표");
            if (rowSummary < 0)
                return;

            var defTaxRaw = ws.Cell(rowSummary + 1, 15).GetString()?.Trim(); // a (O101)
            var diffCarryRaw = ws.Cell(rowSummary + 2, 15).GetString()?.Trim(); // b (O102)
            var globeValRaw = ws.Cell(rowSummary + 3, 15).GetString()?.Trim(); // c (O103)
            var befRecastRaw = ws.Cell(rowSummary + 4, 15).GetString()?.Trim(); // d (O104)
            var totalAdjRaw = ws.Cell(rowSummary + 5, 15).GetString()?.Trim(); // 2. 총조정금액 (O105)
            var preRecastRaw = ws.Cell(rowSummary + 6, 15).GetString()?.Trim(); // e (O106)
            var recastLoRaw = ws.Cell(rowSummary + 7, 15).GetString()?.Trim(); // f (낮은) (O107)
            var recastHiRaw = ws.Cell(rowSummary + 8, 15).GetString()?.Trim(); // g (높은) (O108)
            var totalRaw = ws.Cell(rowSummary + 9, 15).GetString()?.Trim(); // 4. 총이연법인세조정금액 (O109)

            // (b) 조정내역 헤더 탐색 → GIR2501~GIR2516 순서로 16개
            var rowDetail = FindRow(ws, "(b) 조정내역");
            bool hasDetail = false;
            string[] detailAmts = null;
            if (rowDetail >= 0)
            {
                // +1: "1. 이연법인세비용의 조정" 레이블행 (O="순 조정금액") → 스킵
                // +2부터: (a)~(p) 16개 실제 데이터
                detailAmts = new string[DeferTaxAdjItems.Length];
                for (int i = 0; i < DeferTaxAdjItems.Length; i++)
                    detailAmts[i] = ws.Cell(rowDetail + 2 + i, 15).GetString()?.Trim();
                hasDetail = detailAmts.Any(v => !string.IsNullOrEmpty(v));
            }

            bool hasData =
                !string.IsNullOrEmpty(defTaxRaw)
                || !string.IsNullOrEmpty(diffCarryRaw)
                || !string.IsNullOrEmpty(globeValRaw)
                || !string.IsNullOrEmpty(befRecastRaw)
                || !string.IsNullOrEmpty(totalAdjRaw)
                || !string.IsNullOrEmpty(preRecastRaw)
                || !string.IsNullOrEmpty(recastLoRaw)
                || !string.IsNullOrEmpty(recastHiRaw)
                || !string.IsNullOrEmpty(totalRaw)
                || hasDetail;
            if (!hasData)
                return;

            if (etr.EtrStatus == null)
                etr.EtrStatus = new Globe.EtrTypeEtrStatus();
            if (etr.EtrStatus.EtrComputation == null)
                etr.EtrStatus.EtrComputation = new Globe.EtrComputationType();

            var overall =
                etr.EtrStatus.EtrComputation.OverallComputation
                ?? new Globe.EtrComputationTypeOverallComputation();
            etr.EtrStatus.EtrComputation.OverallComputation = overall;

            overall.AdjustedCoveredTax ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTax();
            overall.AdjustedCoveredTax.DeferTaxAdjustAmt ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmt();
            var dt = overall.AdjustedCoveredTax.DeferTaxAdjustAmt;

            // (a) 요약표 scalar 필드
            if (!string.IsNullOrEmpty(defTaxRaw))
                dt.DefTaxAmt = defTaxRaw;
            if (!string.IsNullOrEmpty(diffCarryRaw))
                dt.DiffCarryValue = diffCarryRaw;
            if (!string.IsNullOrEmpty(globeValRaw))
                dt.GLoBeValue = globeValRaw;
            if (!string.IsNullOrEmpty(befRecastRaw))
                dt.BefRecastAdjust = befRecastRaw;
            if (!string.IsNullOrEmpty(totalAdjRaw))
                dt.TotalAdjust = totalAdjRaw;
            if (!string.IsNullOrEmpty(preRecastRaw))
                dt.PreRecast = preRecastRaw;
            if (!string.IsNullOrEmpty(totalRaw))
                dt.Total = totalRaw;

            if (!string.IsNullOrEmpty(recastLoRaw) || !string.IsNullOrEmpty(recastHiRaw))
            {
                dt.Recast ??=
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtRecast();
                if (!string.IsNullOrEmpty(recastLoRaw))
                    dt.Recast.Lower = recastLoRaw;
                if (!string.IsNullOrEmpty(recastHiRaw))
                    dt.Recast.Higher = recastHiRaw;
            }

            // (b) 조정내역 → Adjustments[] (GIR2501~GIR2516)
            if (hasDetail)
            {
                for (int i = 0; i < DeferTaxAdjItems.Length; i++)
                {
                    if (string.IsNullOrEmpty(detailAmts[i]))
                        continue;
                    dt.Adjustments.Add(
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustments
                        {
                            Amount = detailAmts[i],
                            AdjustmentItem = DeferTaxAdjItems[i],
                        }
                    );
                }
            }
        }

        // ─── 3.2.3.1 국가별 선택 → EtrType.Election ─────────────────────────
        private void Map323Election(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            // 3.2.3.1 섹션 헤더 이후부터 검색 (블록 앞쪽 조정사항과 키워드 충돌 방지)
            int rSec = FindRow(ws, "3.2.3.1");
            if (rSec < 0)
                rSec = FindRow(ws, "3.2.3");
            if (rSec < 0)
                return;

            // 헬퍼: B열에서 헤더 찾고 M열 값 반환
            DateTime? ParseYear(string raw, string label)
            {
                if (string.IsNullOrEmpty(raw))
                    return null;
                if (int.TryParse(raw, out var y))
                    return new DateTime(y, 1, 1);
                if (DateTime.TryParse(raw, out var d))
                    return d;
                errors.Add($"[{fileName}] [3.2.3] {label} 연도 파싱 실패: '{raw}'");
                return null;
            }

            // ── 1. 매년 선택 (M열 여/부 bool) — 3.2.3.1 섹션 내부에서 찾기 ──
            int rA = FindRow(ws, "총자산처분이익", rSec);
            int rB = FindRow(ws, "경미한 감액", rSec);
            int rC = FindRow(ws, "실질기반제외소득 미적용", rSec);
            int rD = FindRow(ws, "이월조정대상조세 처리", rSec);

            bool hasAnnual = rA >= 0 || rB >= 0 || rC >= 0 || rD >= 0;

            // ── 2. 5년 선택 (M=선택사업연도 col13, P=취소사업연도 col16) ───
            int rE = FindRow(ws, "지분투자손익포함", rSec);
            int rF = FindRow(ws, "주식기준보상비용", rSec);
            int rG = FindRow(ws, "실현주의 적용", rSec);
            int rH = FindRow(ws, "구성기업 간 연결회계조정", rSec);
            int rI = FindRow(ws, "이연법인세비용 미배분", rSec);

            // ── 5. 그 밖의 선택 (M=선택사업연도 col13, P=취소사업연도 col16) ─
            int rJ = FindRow(ws, "결손취급특례", rSec);

            bool hasAny =
                hasAnnual || rE >= 0 || rF >= 0 || rG >= 0 || rH >= 0 || rI >= 0 || rJ >= 0;
            if (!hasAny)
                return;

            if (etr.EtrStatus == null)
                etr.EtrStatus = new Globe.EtrTypeEtrStatus();
            etr.Election ??= new Globe.EtrTypeElection();
            var el = etr.Election;

            // ── 매년 선택 ──────────────────────────────────────────────────
            if (rA >= 0)
            {
                var v = ws.Cell(rA, 13).GetString()?.Trim(); // M47
                if (!string.IsNullOrEmpty(v))
                {
                    el.Art326 = ParseBool(v);
                    el.Art326Specified = true;
                }
            }
            if (rB >= 0)
            {
                var v = ws.Cell(rB, 13).GetString()?.Trim(); // M175
                if (!string.IsNullOrEmpty(v))
                {
                    el.Art461 = ParseBool(v);
                    el.Art461Specified = true;
                }
            }
            if (rC >= 0)
            {
                var v = ws.Cell(rC, 13).GetString()?.Trim(); // M176
                if (!string.IsNullOrEmpty(v))
                {
                    el.Art531 = ParseBool(v);
                    el.Art531Specified = true;
                }
            }
            if (rD >= 0)
            {
                var v = ws.Cell(rD, 13).GetString()?.Trim(); // M177
                if (!string.IsNullOrEmpty(v))
                {
                    el.Art415 = ParseBool(v);
                    el.Art415Specified = true;
                }
            }

            // ── 5년 선택 공통 헬퍼 ─────────────────────────────────────────
            (DateTime elYear, bool hasEl, DateTime revYear, bool hasRev) ReadElection(int row)
            {
                var mRaw = ws.Cell(row, 13).GetString()?.Trim();
                var pRaw = ws.Cell(row, 16).GetString()?.Trim();
                var ey = ParseYear(mRaw, $"행{row} 선택사업연도");
                var ry = ParseYear(pRaw, $"행{row} 취소사업연도");
                return (ey ?? default, ey.HasValue, ry ?? default, ry.HasValue);
            }

            // e. Art3.2.1c — ElectionYear/RevocationYear 먼저 설정, (b) 필요정보는 아래서 채움
            if (rE >= 0)
            {
                var (ey, hasEl, ry, hasRev) = ReadElection(rE);
                if (hasEl)
                {
                    el.Art321C = new Globe.EtrTypeElectionArt321C
                    {
                        Status = true,
                        ElectionYear = ey,
                        RevocationYear = ry,
                        RevocationYearSpecified = hasRev,
                    };
                }
            }

            // (b) 국가별 선택 관련 필요정보 → Art321C 추가 필드 (O열=15)
            var rowB321C = FindRow(ws, "(b) 국가별 선택과 관련된 필요정보");
            if (rowB321C >= 0 && el.Art321C != null)
            {
                var v1 = ws.Cell(rowB321C + 1, 15).GetString()?.Trim(); // O190
                var v2 = ws.Cell(rowB321C + 2, 15).GetString()?.Trim(); // O191
                var v3 = ws.Cell(rowB321C + 3, 15).GetString()?.Trim(); // O192
                var v4 = ws.Cell(rowB321C + 4, 15).GetString()?.Trim(); // O193
                var v5 = ws.Cell(rowB321C + 5, 15).GetString()?.Trim(); // O194
                if (!string.IsNullOrEmpty(v1))
                    el.Art321C.KEquityInvestmentInclusionElection = v1;
                if (!string.IsNullOrEmpty(v2))
                    el.Art321C.QualOwnerIntentBalance = v2;
                if (!string.IsNullOrEmpty(v3))
                    el.Art321C.Additions = v3;
                if (!string.IsNullOrEmpty(v4))
                    el.Art321C.Reductions = v4;
                if (!string.IsNullOrEmpty(v5))
                    el.Art321C.OutstandingBalance = v5;
            }

            // f. Art3.2.2
            if (rF >= 0)
            {
                var (ey, hasEl, ry, hasRev) = ReadElection(rF);
                if (hasEl)
                    el.Art322 = new Globe.EtrTypeElectionArt322
                    {
                        Status = true,
                        ElectionYear = ey,
                        RevocationYear = ry,
                        RevocationYearSpecified = hasRev,
                    };
            }

            // g. Art3.2.5
            if (rG >= 0)
            {
                var (ey, hasEl, ry, hasRev) = ReadElection(rG);
                if (hasEl)
                    el.Art325 = new Globe.EtrTypeElectionArt325
                    {
                        Status = true,
                        ElectionYear = ey,
                        ElectionYearSpecified = true,
                        RevocationYear = ry,
                        RevocationYearSpecified = hasRev,
                    };
            }

            // h. Art3.2.8
            if (rH >= 0)
            {
                var (ey, hasEl, ry, hasRev) = ReadElection(rH);
                if (hasEl)
                    el.Art328 = new Globe.EtrTypeElectionArt328
                    {
                        Status = true,
                        ElectionYear = ey,
                        ElectionYearSpecified = true,
                        RevocationYear = ry,
                        RevocationYearSpecified = hasRev,
                    };
            }

            // i. NoDefTaxAllocation
            if (rI >= 0)
            {
                var (ey, hasEl, ry, hasRev) = ReadElection(rI);
                if (hasEl)
                    el.NoDefTaxAllocation = new Globe.EtrTypeElectionNoDefTaxAllocation
                    {
                        Status = true,
                        ElectionYear = ey,
                        RevocationYear = ry,
                        RevocationYearSpecified = hasRev,
                    };
            }

            // j. Art4.5
            if (rJ >= 0)
            {
                var (ey, hasEl, ry, hasRev) = ReadElection(rJ);
                if (hasEl)
                    el.Art45 = new Globe.EtrTypeElectionArt45
                    {
                        Status = true,
                        ElectionYear = ey,
                        ElectionYearSpecified = true,
                        RevocationYear = ry,
                        RevocationYearSpecified = hasRev,
                    };
            }

            // 후처리: Election이 실제로 아무 필드도 채워지지 않았으면 null
            // (헤더만 있고 실제 입력값 0개여도 위에서 etr.Election 객체는 생성됨)
            if (etr.Election is { } e0)
            {
                bool elHasAny =
                    e0.Art326Specified
                    || e0.Art415Specified
                    || e0.Art461Specified
                    || e0.Art531Specified
                    || e0.SimplifiedReportingSpecified
                    || e0.Art322 != null
                    || e0.Art325 != null
                    || e0.Art328 != null
                    || e0.NoDefTaxAllocation != null
                    || e0.Art45 != null
                    || e0.Art321C != null;
                if (!elHasAny)
                    etr.Election = null;
            }
        }

        // ─── 3.2.3.2.1 간주분배세액 → AdjustedCoveredTax.DeemedDistTax ─────
        private void Map323DeemedDistTax(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            // 3.2.3.2 헤더 확인
            var row3232 = FindRow(ws, "3.2.3.2");
            if (row3232 < 0)
                return;

            // 간주분배세액 선택 여부: O열 ■ = 선택
            var rowElect = FindRow(ws, "간주분배세액 선택", row3232);
            bool elected = false;
            if (rowElect >= 0)
            {
                var v = ws.Cell(rowElect, 15).GetString()?.Trim(); // O197
                elected = v == "■" || ParseBool(v);
            }

            // (a) 환입계정 운용 표: 헤더 아래 B열=연도(int), C=StartAmount,
            //   F=DdtYear3, I=DdtYear2, K=DdtYear1, M=DdtYear0, P=EndAmount
            var rowA = FindRow(ws, "(a)", row3232);
            var recaptures =
                new List<Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeemedDistTaxElectionRecapture>();
            if (rowA >= 0)
            {
                // 열 헤더 2줄(+1,+2) 건너뛰기. +3부터 스캔.
                // B열: 실제 연도(숫자)를 입력해야 처리. "직전N년차" 등 텍스트는 스킵.
                // "해당없음" → 해당 연도에 환입계정 없음을 의미 → null 처리
                static string FilterNA(string s) =>
                    string.IsNullOrEmpty(s) || s == "해당없음" ? null : s;

                for (int row = rowA + 3; ; row++)
                {
                    var bRaw = ws.Cell(row, 2).GetString()?.Trim();
                    if (string.IsNullOrEmpty(bRaw))
                        break;

                    // B열: 실제 연도(숫자/날짜) 또는 템플릿 라벨("직전 N년차" / "신고대상 사업연도") 모두 처리
                    DateTime recYear;
                    if (int.TryParse(bRaw, out var yearNum))
                        recYear = new DateTime(yearNum, 1, 1);
                    else if (DateTime.TryParse(bRaw, out var yearDate))
                        recYear = yearDate;
                    else if (TryMapLabelToYear(bRaw, out var labelYear))
                        recYear = labelYear;
                    else
                        continue;

                    var cRaw = FilterNA(ws.Cell(row, 3).GetString()?.Trim()); // StartAmount
                    var fRaw = FilterNA(ws.Cell(row, 6).GetString()?.Trim()); // DdtYear3
                    var iRaw = FilterNA(ws.Cell(row, 9).GetString()?.Trim()); // DdtYear2
                    var kRaw = FilterNA(ws.Cell(row, 11).GetString()?.Trim()); // DdtYear1
                    var mRaw = FilterNA(ws.Cell(row, 13).GetString()?.Trim()); // DdtYear0
                    var pRaw = FilterNA(ws.Cell(row, 16).GetString()?.Trim()); // EndAmount

                    bool hasValues =
                        cRaw != null
                        || fRaw != null
                        || iRaw != null
                        || kRaw != null
                        || mRaw != null
                        || pRaw != null;
                    if (!hasValues)
                        continue;

                    // XSD required string 필드 — 빈 셀/해당없음(null)은 "0" 기본
                    var rec =
                        new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeemedDistTaxElectionRecapture
                        {
                            Year = recYear,
                            StartAmount = cRaw ?? "0",
                            DdtYear3 = fRaw ?? "0",
                            DdtYear2 = iRaw ?? "0",
                            DdtYear1 = kRaw ?? "0",
                            DdtYear0 = mRaw ?? "0",
                            TotalDdt = "0", // 폼에 직접 입력칸 없음 — 0 기본
                            EndAmount = pRaw ?? "0",
                        };
                    recaptures.Add(rec);
                }
            }

            // (b) 이탈 등 적용: 헤더 아래 데이터 행 1개
            //   B(2)=Reduction, G(7)=IncrementalTopUpTax, M(13)=Ratio
            var rowBb = FindRow(ws, "(b) 간주분배세액 선택 구성기업", row3232);
            string reductionRaw = null,
                incrRaw = null,
                ratioRaw = null;
            if (rowBb >= 0)
            {
                // +1: 컬럼헤더 행 ("1. 이전 사업연도 조정대상조세 감액" 등) → 스킵
                // +2: 데이터 행 — B(2)=Reduction, G(7)=IncrementalTopUpTax, M(13)=Ratio
                reductionRaw = ws.Cell(rowBb + 2, 2).GetString()?.Trim(); // B210
                incrRaw = ws.Cell(rowBb + 2, 7).GetString()?.Trim(); // G210
                ratioRaw = ws.Cell(rowBb + 2, 13).GetString()?.Trim(); // M210
            }

            bool hasData =
                elected
                || recaptures.Count > 0
                || !string.IsNullOrEmpty(reductionRaw)
                || !string.IsNullOrEmpty(incrRaw)
                || !string.IsNullOrEmpty(ratioRaw);
            if (!hasData)
                return;

            if (etr.EtrStatus?.EtrComputation?.OverallComputation == null)
                return;
            var overall = etr.EtrStatus.EtrComputation.OverallComputation;
            overall.AdjustedCoveredTax ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTax();
            overall.AdjustedCoveredTax.DeemedDistTax ??=
                new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeemedDistTax
                {
                    Total = "",
                };
            var ddt = overall.AdjustedCoveredTax.DeemedDistTax;

            if (
                elected
                || recaptures.Count > 0
                || !string.IsNullOrEmpty(reductionRaw)
                || !string.IsNullOrEmpty(incrRaw)
                || !string.IsNullOrEmpty(ratioRaw)
            )
            {
                ddt.Election ??=
                    new Globe.EtrComputationTypeOverallComputationAdjustedCoveredTaxDeemedDistTaxElection();
                var el = ddt.Election;

                foreach (var rec in recaptures)
                    el.Recapture.Add(rec);

                if (!string.IsNullOrEmpty(reductionRaw))
                    el.Reduction = RoundToInteger(reductionRaw);
                if (!string.IsNullOrEmpty(incrRaw))
                    el.IncrementalTopUpTax = RoundToInteger(incrRaw);
                var ratio = ParsePercentage(ratioRaw);
                if (ratio.HasValue)
                    el.Ratio = ratio.Value;
                else if (!string.IsNullOrEmpty(ratioRaw))
                    errors.Add(
                        $"[{fileName}] [3.2.3.2(b)] 처분환입비율 형식 오류: '{ratioRaw}' (0~1 사이 decimal 입력)"
                    );
            }
        }

        // ─── 3.3 추가세액 계산 → OverallComputation 나머지 필드 ─────────────
        // 국가별 계산 시트 내부의 "3.3 추가세액 계산" 섹션에서 읽음.
        private void Map33TopUpTaxCalc(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            // OverallComputation 확보 (Map321이 이미 만들었을 수 있음)
            if (etr.EtrStatus == null)
                etr.EtrStatus = new Globe.EtrTypeEtrStatus();
            if (etr.EtrStatus.EtrComputation == null)
                etr.EtrStatus.EtrComputation = new Globe.EtrComputationType();
            var overall =
                etr.EtrStatus.EtrComputation.OverallComputation
                ?? new Globe.EtrComputationTypeOverallComputation();
            etr.EtrStatus.EtrComputation.OverallComputation = overall;

            Map331Summary(ws, overall, errors, fileName);
            Map332SubstanceExclusion(ws, overall, errors, fileName);
            Map333AdditionalTopUpTax(ws, overall, errors, fileName);
            Map334Qdmtt(ws, overall, errors, fileName);
        }

        // ─── 3.3.1 추가세액 요약행 ───────────────────────────────────────────
        // 헤더: "3.3.1 추가세액" (+0) → 컬럼헤더(+1) → 데이터행(+2)
        // B(2)=추가세액비율, F(6)=초과이익, P(16)=추가세액
        private void Map331Summary(
            IXLWorksheet ws,
            Globe.EtrComputationTypeOverallComputation overall,
            List<string> errors,
            string fileName
        )
        {
            var rowHdr = FindRow(ws, "3.3.1 추가세액");
            if (rowHdr < 0)
                return;
            var dr = rowHdr + 2; // 데이터 행

            var pctRaw = ws.Cell(dr, 2).GetString()?.Trim(); // B: 추가세액비율
            var excessRaw = ws.Cell(dr, 6).GetString()?.Trim(); // F: 초과이익
            var taxRaw = ws.Cell(dr, 16).GetString()?.Trim(); // P: 추가세액

            var topUpPct = ParsePercentage(pctRaw);
            if (topUpPct.HasValue)
                overall.TopUpTaxPercentage = topUpPct.Value;
            else if (!string.IsNullOrEmpty(pctRaw))
                errors.Add($"[{fileName}] [3.3.1] 추가세액비율 형식 오류: '{pctRaw}' (0~1 사이 decimal 입력, 예: 0.05)");
            if (!string.IsNullOrEmpty(excessRaw))
                overall.ExcessProfits = RoundToInteger(excessRaw);
            if (!string.IsNullOrEmpty(taxRaw))
                overall.TopUpTax = RoundToInteger(taxRaw);
        }

        // ─── 3.3.2.1 실질기반제외소득 합계 → SubstanceExclusion ──────────────
        // 헤더: "3.3.2.1 실질기반제외소득 합계" (+0)
        //   +1: 카테고리 헤더행 (인건비 제외금액 | N=합계)
        //   +2: 컬럼 헤더행 (B=1.인건비, D=2.반영비율, G=3.유형자산, K=4.반영비율, N=5.합계)
        //   +3: 데이터행
        // PEAllocation(3.3.2.2), FTEAllocation(3.3.2.3)은 향후 추가
        private void Map332SubstanceExclusion(
            IXLWorksheet ws,
            Globe.EtrComputationTypeOverallComputation overall,
            List<string> errors,
            string fileName
        )
        {
            // ── 3.3.2.1 실질기반제외소득 합계 ────────────────────────────────
            // 헤더(+0) → 카테고리헤더(+1) → 컬럼헤더(+2) → 데이터행(+3)
            // B(2)=인건비, D(4)=인건비반영비율, H(8)=유형자산장부가액, K(11)=유형자산반영비율, N(14)=합계
            var rowHdr = FindRow(ws, "3.3.2.1 실질기반제외소득 합계");
            if (rowHdr >= 0)
            {
                var dr = rowHdr + 3;

                var total = ws.Cell(dr, 14).GetString()?.Trim(); // N
                var payrollCost = ws.Cell(dr, 2).GetString()?.Trim(); // B
                var payrollMupRaw = ws.Cell(dr, 4).GetString()?.Trim(); // D
                var tangibleVal = ws.Cell(dr, 8).GetString()?.Trim(); // H
                var tangibleMupRaw = ws.Cell(dr, 11).GetString()?.Trim(); // K

                bool hasData =
                    !string.IsNullOrEmpty(total)
                    || !string.IsNullOrEmpty(payrollCost)
                    || !string.IsNullOrEmpty(tangibleVal);
                if (hasData)
                {
                    overall.SubstanceExclusion ??=
                        new Globe.EtrComputationTypeOverallComputationSubstanceExclusion();
                    var se = overall.SubstanceExclusion;
                    if (!string.IsNullOrEmpty(total))
                        se.Total = RoundToInteger(total);
                    if (!string.IsNullOrEmpty(payrollCost))
                        se.PayrollCost = RoundToInteger(payrollCost);
                    if (!string.IsNullOrEmpty(tangibleVal))
                        se.TangibleAssetValue = RoundToInteger(tangibleVal);
                    var pMup = ParsePercentage(payrollMupRaw);
                    if (pMup.HasValue)
                        se.PayrollMarkUp = pMup.Value;
                    else if (!string.IsNullOrEmpty(payrollMupRaw))
                        errors.Add(
                            $"[{fileName}] [3.3.2.1] 인건비 반영비율 형식 오류: '{payrollMupRaw}' (0~1 사이 decimal 입력, 예: 0.05)"
                        );

                    var tMup = ParsePercentage(tangibleMupRaw);
                    if (tMup.HasValue)
                        se.TangibleAssetMarkup = tMup.Value;
                    else if (!string.IsNullOrEmpty(tangibleMupRaw))
                        errors.Add(
                            $"[{fileName}] [3.3.2.1] 유형자산 반영비율 형식 오류: '{tangibleMupRaw}' (0~1 사이 decimal 입력, 예: 0.05)"
                        );
                }
            }

            // ── 3.3.2.2 PE 배분 ──────────────────────────────────────────────
            // 헤더(+0) → 컬럼헤더(+1) → 데이터행(+2~, B비면 종료)
            // B(2)=인건비Total, D(4)=유형자산Total, H(8)=고정사업장소재지국, K(11)=인건비Alloc, N(14)=유형자산Alloc
            // 실제 데이터 첫 행이 있을 때만 SubstanceExclusion 생성 (빈 객체 방지)
            var row3322 = FindRow(ws, "3.3.2.2");
            if (row3322 >= 0)
            {
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
                for (int r = row3322 + 2; r <= lastRow; r++)
                {
                    var payRaw = ws.Cell(r, 2).GetString()?.Trim(); // B
                    if (string.IsNullOrEmpty(payRaw))
                        break;

                    overall.SubstanceExclusion ??=
                        new Globe.EtrComputationTypeOverallComputationSubstanceExclusion();
                    var tanRaw = ws.Cell(r, 4).GetString()?.Trim(); // D
                    var jurRaw = ws.Cell(r, 8).GetString()?.Trim(); // H
                    var paAlRaw = ws.Cell(r, 11).GetString()?.Trim(); // K
                    var taAlRaw = ws.Cell(r, 14).GetString()?.Trim(); // N

                    var jur =
                        new Globe.EtrComputationTypeOverallComputationSubstanceExclusionPeAllocationJurOfOwners();
                    if (!string.IsNullOrEmpty(jurRaw))
                    {
                        if (TryParseEnum<Globe.CountryCodeType>(jurRaw, out var cc))
                        {
                            jur.ResCountryCode = cc;
                            jur.ResCountryCodeSpecified = true;
                        }
                        else
                            errors.Add(
                                $"[{fileName}] [3.3.2.2] 행{r} 국가코드 파싱 실패: '{jurRaw}'"
                            );
                    }
                    overall.SubstanceExclusion.PeAllocation.Add(
                        new Globe.EtrComputationTypeOverallComputationSubstanceExclusionPeAllocation
                        {
                            JurOfOwners = jur,
                            PayrollCost =
                                new Globe.EtrComputationTypeOverallComputationSubstanceExclusionPeAllocationPayrollCost
                                {
                                    Total = payRaw,
                                    Allocation = paAlRaw ?? "0",
                                },
                            TangibleAssetValue =
                                new Globe.EtrComputationTypeOverallComputationSubstanceExclusionPeAllocationTangibleAssetValue
                                {
                                    Total = tanRaw ?? "0",
                                    Allocation = taAlRaw ?? "0",
                                },
                        }
                    );
                }
            }

            // ── 3.3.2.3 FTE 배분 ─────────────────────────────────────────────
            // 동일 구조: B=인건비Total, D=유형자산Total, H=주주구성기업소재지국, K=인건비Alloc, N=유형자산Alloc
            // 실제 데이터 첫 행이 있을 때만 SubstanceExclusion 생성 (빈 객체 방지)
            var row3323 = FindRow(ws, "3.3.2.3");
            if (row3323 >= 0)
            {
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
                for (int r = row3323 + 2; r <= lastRow; r++)
                {
                    var payRaw = ws.Cell(r, 2).GetString()?.Trim(); // B
                    if (string.IsNullOrEmpty(payRaw))
                        break;
                    var tanRaw = ws.Cell(r, 4).GetString()?.Trim(); // D
                    var jurRaw = ws.Cell(r, 8).GetString()?.Trim(); // H
                    var paAlRaw = ws.Cell(r, 11).GetString()?.Trim(); // K
                    var taAlRaw = ws.Cell(r, 14).GetString()?.Trim(); // N

                    overall.SubstanceExclusion ??=
                        new Globe.EtrComputationTypeOverallComputationSubstanceExclusion();

                    var jur =
                        new Globe.EtrComputationTypeOverallComputationSubstanceExclusionFteAllocationJurOfOwners();
                    if (!string.IsNullOrEmpty(jurRaw))
                    {
                        if (TryParseEnum<Globe.CountryCodeType>(jurRaw, out var cc))
                        {
                            jur.ResCountryCode = cc;
                            jur.ResCountryCodeSpecified = true;
                        }
                        else
                            errors.Add(
                                $"[{fileName}] [3.3.2.3] 행{r} 국가코드 파싱 실패: '{jurRaw}'"
                            );
                    }
                    overall.SubstanceExclusion.FteAllocation.Add(
                        new Globe.EtrComputationTypeOverallComputationSubstanceExclusionFteAllocation
                        {
                            JurOfOwners = jur,
                            PayrollCost =
                                new Globe.EtrComputationTypeOverallComputationSubstanceExclusionFteAllocationPayrollCost
                                {
                                    Total = payRaw,
                                    Allocation = paAlRaw ?? "0",
                                },
                            TangibleAssetValue =
                                new Globe.EtrComputationTypeOverallComputationSubstanceExclusionFteAllocationTangibleAssetValue
                                {
                                    Total = tanRaw ?? "0",
                                    Allocation = taAlRaw ?? "0",
                                },
                        }
                    );
                }
            }
        }

        // ─── 3.3.3 당기추가세액가산액 → AdditionalTopUpTax ──────────────────
        // 3.3.3.1 NonArt415: 항목당 2행 (셀병합 B·C = 1행에만 값, r+1이 재계산 행)
        //   B(2)=관련근거(병합), C(3)=관련사업연도(병합)
        //   G(7)=순글로벌최저한세소득결손, J(10)=조정대상조세, K(11)=실효세율,
        //   L(12)=초과이익, N(14)=추가세액비율, O(15)=추가세액, Q(17)=당기추가세액가산액
        // 3.3.3.2 Art415: "1.~4." 항목 → M열(13)
        private void Map333AdditionalTopUpTax(
            IXLWorksheet ws,
            Globe.EtrComputationTypeOverallComputation overall,
            List<string> errors,
            string fileName
        )
        {
            // ── 3.3.3.1 NonArt4.1.5 ──────────────────────────────────────────
            var row3331 = FindRow(ws, "3.3.3.1");
            if (row3331 >= 0)
            {
                // 컬럼헤더(+1), 서브헤더(+2), 데이터행(+3~, 2행씩)
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
                for (int r = row3331 + 3; r + 1 <= lastRow; r += 2)
                {
                    // B·C 셀은 2행 병합 → 첫 번째 행에서만 읽기
                    var articlesRaw = ws.Cell(r, 2).GetString()?.Trim(); // B
                    if (string.IsNullOrEmpty(articlesRaw))
                        break;
                    var yearRaw = ws.Cell(r, 3).GetString()?.Trim(); // C

                    if (!DateTime.TryParse(yearRaw, out var year) && !int.TryParse(yearRaw, out _))
                    {
                        errors.Add($"[{fileName}] [3.3.3.1] 행{r} 사업연도 파싱 실패: '{yearRaw}'");
                        // for 루프의 r += 2가 자동으로 다음 쌍으로 진행 — 추가 보정 불필요
                        continue;
                    }
                    if (int.TryParse(yearRaw, out var yi))
                        year = new DateTime(yi, 12, 31);

                    // row r = 당초신고(Previous), row r+1 = 재계산(Recalculated)
                    static string Cell(IXLWorksheet w, int row, int col) =>
                        w.Cell(row, col).GetString()?.Trim();

                    var entry =
                        new Globe.EtrComputationTypeOverallComputationAdditionalTopUpTaxNonArt415
                        {
                            Year = year,
                            AdditionalTopUpTax = RoundToInteger(Cell(ws, r, 17)) ?? "0", // Q: 당기추가세액가산액
                        };

                    foreach (
                        var code in articlesRaw.Split(
                            ',',
                            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
                        )
                    )
                    {
                        if (TryParseEnum<Globe.NonArt415EnumType>(code, out var art))
                            entry.Articles.Add(art);
                        else
                            errors.Add(
                                $"[{fileName}] [3.3.3.1] 행{r} 관련근거 파싱 실패: '{code}'"
                            );
                    }

                    // Previous (당초신고, 행 r) — XSD required xs:integer 필드는 "0" 기본
                    entry.Previous =
                        new Globe.EtrComputationTypeOverallComputationAdditionalTopUpTaxNonArt415Previous
                        {
                            NetGlobeIncome = RoundToInteger(Cell(ws, r, 7)) ?? "0", // G
                            AdjustedCoveredTax = RoundToInteger(Cell(ws, r, 10)) ?? "0", // J
                            ExcessProfits = RoundToInteger(Cell(ws, r, 12)) ?? "0", // L
                            TopUpTax = RoundToInteger(Cell(ws, r, 15)) ?? "0", // O
                        };
                    var prevEtrRaw = Cell(ws, r, 11); // K
                    var prevPctRaw = Cell(ws, r, 14); // N
                    var pEtr = ParsePercentage(prevEtrRaw);
                    if (pEtr.HasValue) entry.Previous.EtrRate = pEtr.Value;
                    var pPct = ParsePercentage(prevPctRaw);
                    if (pPct.HasValue) entry.Previous.TopUpTaxPercentage = pPct.Value;

                    // Recalculated (재계산, 행 r+1) — XSD required xs:integer 필드는 "0" 기본
                    entry.Recalculated =
                        new Globe.EtrComputationTypeOverallComputationAdditionalTopUpTaxNonArt415Recalculated
                        {
                            NetGlobeIncome = RoundToInteger(Cell(ws, r + 1, 7)) ?? "0", // G
                            AdjustedCoveredTax = RoundToInteger(Cell(ws, r + 1, 10)) ?? "0", // J
                            ExcessProfits = RoundToInteger(Cell(ws, r + 1, 12)) ?? "0", // L
                            TopUpTax = RoundToInteger(Cell(ws, r + 1, 15)) ?? "0", // O
                        };
                    var recEtrRaw = Cell(ws, r + 1, 11);
                    var recPctRaw = Cell(ws, r + 1, 14);
                    var rEtr = ParsePercentage(recEtrRaw);
                    if (rEtr.HasValue) entry.Recalculated.EtrRate = rEtr.Value;
                    var rPct = ParsePercentage(recPctRaw);
                    if (rPct.HasValue) entry.Recalculated.TopUpTaxPercentage = rPct.Value;

                    overall.AdditionalTopUpTax ??=
                        new Globe.EtrComputationTypeOverallComputationAdditionalTopUpTax();
                    overall.AdditionalTopUpTax.NonArt415.Add(entry);
                }
            }

            // ── 3.3.3.2 Art4.1.5 — 1~4 모두 M열(13) ─────────────────────────
            var row3332 = FindRow(ws, "3.3.3.2");
            if (row3332 >= 0)
            {
                var rAct = FindRow(ws, "1. 소재지국 조정대상조세", row3332);
                var rLoss = FindRow(ws, "2. 소재지국 순글로벌최저한세결손", row3332);
                var rExp = FindRow(ws, "3. 조정대상조세 예상액", row3332);
                var rAdd = FindRow(ws, "4. 당기추가세액가산액", row3332);

                var actRaw = rAct >= 0 ? ws.Cell(rAct, 13).GetString()?.Trim() : null; // M (M245)
                var lossRaw = rLoss >= 0 ? ws.Cell(rLoss, 13).GetString()?.Trim() : null; // M246
                var expRaw = rExp >= 0 ? ws.Cell(rExp, 13).GetString()?.Trim() : null; // M247
                var addRaw = rAdd >= 0 ? ws.Cell(rAdd, 13).GetString()?.Trim() : null; // M248

                // Art4.1.5는 NetGlobeIncome이 음수(loss)일 때 보고하는 특례.
                // 모든 값이 0/빈 값이면 보고 의도 없음 — emit 자체 생략.
                // (검증 70089: AdjustedCoveredTax는 음수, 70090: GlobeLoss = NetGlobeIncome.Total)
                bool hasArt415 =
                    HasNonZeroValue(actRaw)
                    || HasNonZeroValue(lossRaw)
                    || HasNonZeroValue(expRaw)
                    || HasNonZeroValue(addRaw);
                if (hasArt415)
                {
                    overall.AdditionalTopUpTax ??=
                        new Globe.EtrComputationTypeOverallComputationAdditionalTopUpTax();
                    // Art4.1.5는 4개 필드 모두 XSD required. 사용자 입력이 일부만 있으면
                    // 누락 필드(빈 셀 또는 null)는 "0" 기본값으로 채워 XSD 유효성 보장.
                    overall.AdditionalTopUpTax.Art415 =
                        new Globe.EtrComputationTypeOverallComputationAdditionalTopUpTaxArt415
                        {
                            AdjustedCoveredTax = NullIfEmpty(RoundToInteger(actRaw)) ?? "0",
                            GlobeLoss = NullIfEmpty(RoundToInteger(lossRaw)) ?? "0",
                            ExpectedAdjustedCoveredTax = NullIfEmpty(RoundToInteger(expRaw)) ?? "0",
                            AdditionalTopUpTax = NullIfEmpty(RoundToInteger(addRaw)) ?? "0",
                        };
                }
            }
        }

        // ─── 3.3.4 적격소재국추가세액 → QDMTT ───────────────────────────────
        // 1~5, 7~8 모두 K열(11)이 값
        // 6. CurrencyElection: 행은 헤더(K=통화라벨), 데이터는 행+1에서 K=통화, M=선택, P=취소
        private void Map334Qdmtt(
            IXLWorksheet ws,
            Globe.EtrComputationTypeOverallComputation overall,
            List<string> errors,
            string fileName
        )
        {
            var rowHdr = FindRow(ws, "3.3.4 적격소재국추가세액");
            if (rowHdr < 0)
                return;

            var rFas = FindRow(ws, "1. 회계기준", rowHdr);
            var rAmt = FindRow(ws, "2. 적격소재국추가세액", rowHdr);
            var rRate = FindRow(ws, "3. 적격소재국추가세제도 최저한세율", rowHdr);
            var rBasis = FindRow(ws, "4. 실효세율 계산을 위한", rowHdr);
            var rCur = FindRow(ws, "5. 사용 통화", rowHdr);
            var rCurEl = FindRow(ws, "6. 연결재무제표", rowHdr);
            var rSbie = FindRow(ws, "7. 실질기반제외소득 적용가능", rowHdr);
            var rDeMin = FindRow(ws, "8. 최소적용제외 특례", rowHdr);

            var fasRaw = rFas >= 0 ? ws.Cell(rFas, 11).GetString()?.Trim() : null; // K (K251)
            var amtRaw = rAmt >= 0 ? ws.Cell(rAmt, 11).GetString()?.Trim() : null; // K252
            var rateRaw = rRate >= 0 ? ws.Cell(rRate, 11).GetString()?.Trim() : null; // K253
            var basisRaw = rBasis >= 0 ? ws.Cell(rBasis, 11).GetString()?.Trim() : null; // K254
            var curRaw = rCur >= 0 ? ws.Cell(rCur, 11).GetString()?.Trim() : null; // K255
            var sbieRaw = rSbie >= 0 ? ws.Cell(rSbie, 11).GetString()?.Trim() : null; // K258
            var deMinRaw = rDeMin >= 0 ? ws.Cell(rDeMin, 11).GetString()?.Trim() : null; // K259

            // CurrencyElection: rCurEl 행 헤더(K=통화라벨, M=선택년도라벨), 데이터는 rCurEl+1
            var rCurElData = rCurEl >= 0 ? rCurEl + 1 : -1;
            var curElCurRaw = rCurElData >= 0 ? ws.Cell(rCurElData, 11).GetString()?.Trim() : null; // K
            var curElSelRaw = rCurElData >= 0 ? ws.Cell(rCurElData, 13).GetString()?.Trim() : null; // M
            var curElRevRaw = rCurElData >= 0 ? ws.Cell(rCurElData, 16).GetString()?.Trim() : null; // P

            bool hasData =
                !string.IsNullOrEmpty(fasRaw)
                || !string.IsNullOrEmpty(amtRaw)
                || !string.IsNullOrEmpty(curRaw);
            if (!hasData)
                return;

            overall.Qdmtt ??= new Globe.EtrComputationTypeOverallComputationQdmtt();
            var q = overall.Qdmtt;

            if (!string.IsNullOrEmpty(fasRaw))
                q.Fas = fasRaw;
            if (!string.IsNullOrEmpty(amtRaw))
                q.Amount = RoundToInteger(amtRaw);
            if (!string.IsNullOrEmpty(basisRaw))
                q.BasisforBlending = basisRaw;

            var minRate = ParsePercentage(rateRaw);
            if (minRate.HasValue)
            {
                q.MinRate = minRate.Value;
                q.MinRateSpecified = true;
            }
            else if (!string.IsNullOrEmpty(rateRaw))
                errors.Add($"[{fileName}] [3.3.4] 최저한세율 형식 오류: '{rateRaw}' (0~1 사이 decimal 입력, 예: 0.15)");

            if (!string.IsNullOrEmpty(curRaw))
            {
                if (TryParseEnum<Globe.CurrCodeType>(curRaw, out var cur))
                    q.Currency = cur;
                else
                    errors.Add($"[{fileName}] [3.3.4] 사용통화 파싱 실패: '{curRaw}'");
            }

            if (!string.IsNullOrEmpty(sbieRaw))
                q.SbieAvailable = ParseBool(sbieRaw);
            if (!string.IsNullOrEmpty(deMinRaw))
                q.DeMinAvailable = ParseBool(deMinRaw);

            // CurrencyElection (있는 경우) — Currency는 CurrencyEnumType
            if (!string.IsNullOrEmpty(curElCurRaw) && TryParseDate(curElSelRaw, out var selY))
            {
                if (TryParseEnum<Globe.CurrencyEnumType>(curElCurRaw, out var elCur))
                {
                    q.CurrencyElection ??=
                        new Globe.EtrComputationTypeOverallComputationQdmttCurrencyElection();
                    q.CurrencyElection.Currency = elCur;
                    q.CurrencyElection.Status = true;
                    q.CurrencyElection.ElectionYear = selY;

                    if (
                        !string.IsNullOrEmpty(curElRevRaw)
                        && TryParseDate(curElRevRaw, out var revY)
                    )
                    {
                        q.CurrencyElection.RevocationYear = revY;
                        q.CurrencyElection.RevocationYearSpecified = true;
                    }
                }
                else
                    errors.Add(
                        $"[{fileName}] [3.3.4] CurrencyElection 통화 파싱 실패: '{curElCurRaw}'"
                    );
            }
        }

        // ─── 3.2.4.4 국제해운소득·결손 제외 → NetGlobeIncome.IntShippingIncome ──────
        // 국가별 계산 시트 내부의 "(b) 적격국제해운부수소득" 섹션에서 읽음.
        private void MapShippingIncome(
            IXLWorksheet ws,
            Globe.EtrType etr,
            List<string> errors,
            string fileName
        )
        {
            // 각 항목을 B열 텍스트로 탐색 (행 추가/이동 대응)
            // 값 열: N열(14)
            var rowHdr = FindRow(ws, "(b) 적격국제해운부수소득");
            if (rowHdr < 0)
                return;

            var r1 = FindRow(ws, "1. 모든 구성기업", rowHdr);
            var r2 = FindRow(ws, "2. 50% 한도", rowHdr);
            var r3 = FindRow(ws, "3. 모든 구성기업", rowHdr);
            var r4 = FindRow(ws, "4. B가 A의 50%", rowHdr);

            var totalIntShipRaw = r1 >= 0 ? ws.Cell(r1, 14).GetString()?.Trim() : null; // N214
            var fiftyCapRaw = r2 >= 0 ? ws.Cell(r2, 14).GetString()?.Trim() : null; // N215
            var totalQualAncRaw = r3 >= 0 ? ws.Cell(r3, 14).GetString()?.Trim() : null; // N216
            var excessCapRaw = r4 >= 0 ? ws.Cell(r4, 14).GetString()?.Trim() : null; // N217

            // 의미있는 해운소득(국제해운/적격부수) 둘 중 하나라도 0 아닌 값이 있을 때만 생성
            // FiftyPercentCap/ExcessOfCap만 "0"으로 채워진 경우는 빈 구조체로 간주
            bool hasData =
                HasNonZeroValue(totalIntShipRaw)
                || HasNonZeroValue(totalQualAncRaw);
            if (!hasData)
                return;

            // OverallComputation.NetGlobeIncome 확보
            if (etr.EtrStatus == null)
                etr.EtrStatus = new Globe.EtrTypeEtrStatus();
            if (etr.EtrStatus.EtrComputation == null)
                etr.EtrStatus.EtrComputation = new Globe.EtrComputationType();
            var overall =
                etr.EtrStatus.EtrComputation.OverallComputation
                ?? new Globe.EtrComputationTypeOverallComputation();
            etr.EtrStatus.EtrComputation.OverallComputation = overall;
            overall.NetGlobeIncome ??=
                new Globe.EtrComputationTypeOverallComputationNetGlobeIncome();

            // Total = TotalIntShipIncome + TotalQualifiedAncIncome - ExcessOfCap (수치 계산)
            string totalRaw = "";
            if (
                decimal.TryParse(totalIntShipRaw, out var a)
                && decimal.TryParse(totalQualAncRaw, out var b)
            )
            {
                var excess = decimal.TryParse(excessCapRaw, out var e) ? e : 0m;
                totalRaw = (a + b - excess).ToString();
            }

            overall.NetGlobeIncome.IntShippingIncome =
                new Globe.EtrComputationTypeOverallComputationNetGlobeIncomeIntShippingIncome
                {
                    Total = NullIfEmpty(totalRaw),
                    TotalIntShipIncome = NullIfEmpty(totalIntShipRaw),
                    FiftyPercentCap = NullIfEmpty(fiftyCapRaw),
                    TotalQualifiedAncIncome = NullIfEmpty(totalQualAncRaw),
                    ExcessOfCap = NullIfEmpty(excessCapRaw),
                };
        }
    }
}
