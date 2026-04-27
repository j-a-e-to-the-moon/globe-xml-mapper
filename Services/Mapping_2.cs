using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 시트 2: 국가별 적용면제 및 제외.
    /// 블록1(b1~b1+20) + 간격(2행) + 블록2(b2~b2+28) = 52행 세트.
    /// 다른 매퍼와 동일하게 B열 헤더("2.1 국가별 기본사항") 탐색으로 다중 블록 처리.
    ///
    /// 검증된 행 오프셋 (b1=2, b2=25 기준):
    ///   2.1  : O(b1+3)=국가, O(b1+6)=과세권국가
    ///   2.2.1: O(b1+12)=SafeHarbour
    ///   2.2.1.2: H/N(b1+16~b1+19)=간소화계산
    ///   2.2.1.3: O(b2+0)=수익, O(b2+1)=세전손익, O(b2+2)=간이세액, O(b2+5)=CIT율
    ///   2.2.2: B(b2+8)=GIR2901체크, B(b2+9)=GIR2902체크
    ///          E/I/L/O(b2+11~b2+14)=FinancialData(신고/직전/직전전/평균)
    ///   2.3  : M(b2+17)=개시일, M(b2+18)=준거국가, M(b2+19)=유형자산,
    ///          M(b2+20)=국가수, M(b2+21)=준거국가 외 유형자산(통합 셀 — 국가,값; 국가,값)
    /// </summary>
    public class Mapping_2 : MappingBase
    {
        private const string BLOCK_ANCHOR = "2.1 국가별 기본사항"; // b1+2 위치의 헤더
        private const int ANCHOR_TO_B1_OFFSET = 2; // b1 = 앵커행 - 2
        private const int BLOCK1_SIZE = 21; // 21행 (b1 ~ b1+20)
        private const int GAP = 2; // b1+21, b1+22 (블록1과 블록2 사이)

        public Mapping_2()
            : base(null) { }

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            var blockStarts = FindAllBlockStarts(ws);
            for (int idx = 0; idx < blockStarts.Count; idx++)
            {
                var b1 = blockStarts[idx];
                var b2 = b1 + BLOCK1_SIZE + GAP;
                MapOneCountry(ws, globe, errors, fileName, b1, b2, idx + 1);
            }
        }

        // B열에서 BLOCK_ANCHOR 텍스트를 포함하는 행을 모두 찾아 b1으로 변환.
        // 각 적용면제 블록은 b1=앵커행-2 위치에서 시작.
        private static List<int> FindAllBlockStarts(IXLWorksheet ws)
        {
            var result = new List<int>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 200;
            for (int r = 1; r <= lastRow; r++)
            {
                var v = ws.Cell(r, 2).GetString()?.Trim() ?? "";
                if (v.Contains(BLOCK_ANCHOR))
                    result.Add(r - ANCHOR_TO_B1_OFFSET);
            }
            return result;
        }

        private void MapOneCountry(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName,
            int b1,
            int b2,
            int blockNum
        )
        {
            // ─── 2.1 소재지국 (O5 = b1+3) ─────────────────────────────────
            var jurCode = ws.Cell(b1 + 3, 15).GetString()?.Trim(); // O5
            if (string.IsNullOrEmpty(jurCode))
                return;

            if (!TryParseEnum<Globe.CountryCodeType>(jurCode, out var countryCode))
            {
                errors.Add(
                    $"[{fileName}] 적용면제 블록{blockNum}: 소재지국 코드 '{jurCode}' 파싱 실패"
                );
                return;
            }

            var loc = $"2 적용면제/블록{blockNum}('{jurCode}')";

            // ─── Summary 찾기 또는 생성 ───────────────────────────────────
            var summary = globe.GlobeBody.Summary.FirstOrDefault(s =>
                s.Jurisdiction?.JurisdictionNameSpecified == true
                && s.Jurisdiction.JurisdictionName == countryCode
            );
            if (summary == null)
            {
                summary = new Globe.GlobeBodyTypeSummary
                {
                    Jurisdiction = new Globe.SummaryTypeJurisdiction
                    {
                        JurisdictionName = countryCode,
                        JurisdictionNameSpecified = true,
                    },
                };
                globe.GlobeBody.Summary.Add(summary);
            }

            // 과세권 국가 (O8=b1+6) — JurisdictionSection.JurWithTaxingRights로 이동 (js 생성 후 처리)
            var taxJurRaw = ws.Cell(b1 + 6, 15).GetString()?.Trim(); // O8

            // ─── 2.2.1 적용면제 (O14=b1+12) ──────────────────────────────
            var safeHarbourRaw = ws.Cell(b1 + 12, 15).GetString()?.Trim(); // O14
            if (!string.IsNullOrEmpty(safeHarbourRaw))
            {
                foreach (
                    var code in safeHarbourRaw.Split(
                        ',',
                        StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
                    )
                )
                    SetEnum<Globe.SafeHarbourEnumType>(
                        code,
                        v =>
                        {
                            if (!summary.SafeHarbour.Contains(v))
                                summary.SafeHarbour.Add(v);
                        },
                        errors,
                        fileName,
                        new MappingEntry { Cell = $"O{b1 + 12}", Label = $"[{loc}] 적용면제" }
                    );
            }

            // ─── 2.2.1.2 간소화 데이터 (H18-N21 = b1+16 ~ b1+19) ────────
            var s1Rev = ws.Cell(b1 + 16, 8).GetString()?.Trim(); // H18
            var s1Tax = ws.Cell(b1 + 16, 14).GetString()?.Trim(); // N18
            var s2Rev = ws.Cell(b1 + 17, 8).GetString()?.Trim(); // H19
            var s2Tax = ws.Cell(b1 + 17, 14).GetString()?.Trim(); // N19
            var s3Rev = ws.Cell(b1 + 18, 8).GetString()?.Trim(); // H20
            var s3Tax = ws.Cell(b1 + 18, 14).GetString()?.Trim(); // N20
            var saRev = ws.Cell(b1 + 19, 8).GetString()?.Trim(); // H21
            var saTax = ws.Cell(b1 + 19, 14).GetString()?.Trim(); // N21

            // ─── 2.2.1.3 전환기 데이터 (b2+0 ~ b2+5) ─────────────────────
            // O25=b2+0: 총수익, O26=b2+1: 세전손익, O27=b2+2: 간이대상조세
            // O30=b2+5: 법인세 명목세율
            var cbcrRev = ws.Cell(b2 + 0, 15).GetString()?.Trim(); // O25
            var cbcrPl = ws.Cell(b2 + 1, 15).GetString()?.Trim(); // O26
            var cbcrTax = ws.Cell(b2 + 2, 15).GetString()?.Trim(); // O27
            var utprRate = ws.Cell(b2 + 5, 15).GetString()?.Trim(); // O30

            // ─── 2.2.2 체크박스 (b2+8=GIR2901, b2+9=GIR2902) ──────────────
            // R33=b2+8: □/■ 신고대상 사업연도 → GIR2901
            // R34=b2+9: □/■ 중요성이 낮은 구성기업 → GIR2902
            Globe.DeminimisSimpleBasisEnumType? deminiBasis = null;
            if (RowContains(ws, b2 + 8, "■"))
                deminiBasis = Globe.DeminimisSimpleBasisEnumType.Gir2901;
            else if (RowContains(ws, b2 + 9, "■"))
                deminiBasis = Globe.DeminimisSimpleBasisEnumType.Gir2902;

            // ─── 2.2.2 상세 재무 데이터 (b2+11 ~ b2+14) ──────────────────
            // R36=b2+11: 신고대상, R37=b2+12: 직전, R38=b2+13: 직전전, R39=b2+14: 3년평균
            // 열: E(5)=회계매출, I(9)=GloBE매출, L(12)=회계순이익, O(15)=GloBE소득
            var f1AcRev = ws.Cell(b2 + 11, 5).GetString()?.Trim(); // E36
            var f1GbRev = ws.Cell(b2 + 11, 9).GetString()?.Trim(); // I36
            var f1AcPl = ws.Cell(b2 + 11, 12).GetString()?.Trim(); // L36
            var f1GbPl = ws.Cell(b2 + 11, 15).GetString()?.Trim(); // O36
            var f2AcRev = ws.Cell(b2 + 12, 5).GetString()?.Trim(); // E37
            var f2GbRev = ws.Cell(b2 + 12, 9).GetString()?.Trim(); // I37
            var f2AcPl = ws.Cell(b2 + 12, 12).GetString()?.Trim(); // L37
            var f2GbPl = ws.Cell(b2 + 12, 15).GetString()?.Trim(); // O37
            var f3AcRev = ws.Cell(b2 + 13, 5).GetString()?.Trim(); // E38
            var f3GbRev = ws.Cell(b2 + 13, 9).GetString()?.Trim(); // I38
            var f3AcPl = ws.Cell(b2 + 13, 12).GetString()?.Trim(); // L38
            var f3GbPl = ws.Cell(b2 + 13, 15).GetString()?.Trim(); // O38
            var faAcRev = ws.Cell(b2 + 14, 5).GetString()?.Trim(); // E39
            var faGbRev = ws.Cell(b2 + 14, 9).GetString()?.Trim(); // I39
            var faAcPl = ws.Cell(b2 + 14, 12).GetString()?.Trim(); // L39
            var faGbPl = ws.Cell(b2 + 14, 15).GetString()?.Trim(); // O39

            // ─── 2.3 해외진출 초기 특례 (b2+17 ~ b2+21) ──────────────────
            // M42=b2+17: 1.개시일, M43=b2+18: 2.준거국가, M44=b2+19: 3.유형자산,
            // M45=b2+20: 4.국가수, M46=b2+21: 5.준거국가 외 유형자산(통합 셀)
            var initStartRaw = ws.Cell(b2 + 17, 13).GetString()?.Trim(); // M42
            var initRefJur = ws.Cell(b2 + 18, 13).GetString()?.Trim(); // M43
            var initRefAsset = ws.Cell(b2 + 19, 13).GetString()?.Trim(); // M44
            var initNumJur = ws.Cell(b2 + 20, 13).GetString()?.Trim(); // M45
            var initOtherJurRaw = ws.Cell(b2 + 21, 13).GetString()?.Trim(); // M46 (통합 셀)

            // ─── ETR / InitialIntActivity 데이터 유무 판단 ────────────────
            // "0" 값만 입력된 셀은 데이터 없음으로 간주 (HasNonZeroValue).
            // 그렇지 않으면 사용자가 세이프하버 체크박스 안 했는데도 0으로 채워진 셀 때문에
            // Deminimis 구조가 잘못 생성됨.
            bool hasDemini =
                deminiBasis.HasValue
                || HasNonZeroValue(f1GbRev)
                || HasNonZeroValue(f1AcRev)
                || HasNonZeroValue(s1Rev);
            bool hasCbcr =
                HasNonZeroValue(cbcrRev)
                || HasNonZeroValue(cbcrPl)
                || HasNonZeroValue(cbcrTax);
            bool hasUtpr = !string.IsNullOrEmpty(utprRate);
            bool hasEtrData = hasDemini || hasCbcr || hasUtpr;
            bool hasInit = !string.IsNullOrEmpty(initStartRaw);
            bool hasJurWithTaxingRights = !string.IsNullOrEmpty(taxJurRaw);

            if (!hasEtrData && !hasInit && !hasJurWithTaxingRights)
                return;

            // ─── JurisdictionSection 찾기 또는 생성 ──────────────────────
            var js = globe.GlobeBody.JurisdictionSection.FirstOrDefault(s =>
                s.Jurisdiction == countryCode
            );
            if (js == null)
            {
                js = new Globe.GlobeBodyTypeJurisdictionSection();
                js.Jurisdiction = countryCode;
                js.RecJurCode.Add(countryCode);
                js.GLoBeTax = new Globe.GlobeTax();
                globe.GlobeBody.JurisdictionSection.Add(js);
            }

            // ─── 과세권 국가 → JurisdictionSection.JurWithTaxingRights ──────
            // 복수 항목은 세미콜론으로 구분: "KR; JP" 또는 "KR, (GIR1101,...); JP"
            if (hasJurWithTaxingRights)
            {
                foreach (
                    var entry in taxJurRaw.Split(
                        ';',
                        StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
                    )
                )
                {
                    var jwr = ParseJwrEntry(entry, errors, fileName, loc);
                    if (jwr == null)
                        continue;
                    js.JurWithTaxingRights.Add(jwr);
                }
            }

            // ─── ETR 항목 생성 ────────────────────────────────────────────
            if (hasEtrData)
            {
                var etr = new Globe.EtrType { EtrStatus = new Globe.EtrTypeEtrStatus() };
                var exception = new Globe.EtrTypeEtrStatusEtrException();
                etr.EtrStatus.EtrException = exception;

                // DeminimisSimplifiedNmceCalc
                if (hasDemini)
                {
                    var dmCalc = new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalc
                    {
                        Basis = deminiBasis ?? Globe.DeminimisSimpleBasisEnumType.Gir2901,
                    };

                    var periodEnd =
                        globe.GlobeBody.FilingInfo?.Period?.End
                        ?? new DateTime(DateTime.Today.Year, 12, 31);

                    // 2.2.2 상세 데이터 우선, 없으면 2.2.1.2 간소화 데이터
                    bool useFullData =
                        !string.IsNullOrEmpty(f1GbRev) || !string.IsNullOrEmpty(f1AcRev);

                    if (useFullData)
                    {
                        TryAddFinancialData(dmCalc, periodEnd, f1AcRev, f1GbRev, f1GbPl, f1AcPl);
                        TryAddFinancialData(
                            dmCalc,
                            periodEnd.AddYears(-1),
                            f2AcRev,
                            f2GbRev,
                            f2GbPl,
                            f2AcPl
                        );
                        TryAddFinancialData(
                            dmCalc,
                            periodEnd.AddYears(-2),
                            f3AcRev,
                            f3GbRev,
                            f3GbPl,
                            f3AcPl
                        );

                        if (!string.IsNullOrEmpty(faGbRev) || !string.IsNullOrEmpty(faAcRev))
                        {
                            dmCalc.Average =
                                new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcAverage
                                {
                                    Revenue = NullIfEmpty(faAcRev),
                                    GlobeRevenue = NullIfEmpty(faGbRev),
                                    NetGlobeIncome = NullIfEmpty(faGbPl),
                                    Fanil = NullIfEmpty(faAcPl),
                                };
                        }
                    }
                    else
                    {
                        TryAddSimpleFinancialData(dmCalc, periodEnd, s1Rev, s1Tax);
                        TryAddSimpleFinancialData(dmCalc, periodEnd.AddYears(-1), s2Rev, s2Tax);
                        TryAddSimpleFinancialData(dmCalc, periodEnd.AddYears(-2), s3Rev, s3Tax);

                        if (!string.IsNullOrEmpty(saRev) || !string.IsNullOrEmpty(saTax))
                        {
                            // 간소화: Revenue=GlobeRevenue, NetGlobeIncome=FANIL 복제 (XSD required)
                            dmCalc.Average =
                                new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcAverage
                                {
                                    Revenue = NullIfEmpty(saRev),
                                    GlobeRevenue = NullIfEmpty(saRev),
                                    NetGlobeIncome = NullIfEmpty(saTax),
                                    Fanil = NullIfEmpty(saTax),
                                };
                        }
                    }

                    exception.DeminimisSimplifiedNmceCalc = dmCalc;
                }

                // TransitionalCbCrSafeHarbour
                if (hasCbcr)
                {
                    exception.TransitionalCbCrSafeHarbour =
                        new Globe.EtrTypeEtrStatusEtrExceptionTransitionalCbCrSafeHarbour
                        {
                            Revenue = NullIfEmpty(cbcrRev),
                            Profit = NullIfEmpty(cbcrPl),
                            IncomeTax = NullIfEmpty(cbcrTax),
                        };
                }

                // UtprSafeHarbour
                if (hasUtpr)
                {
                    var citRate = ParsePercentage(utprRate);
                    if (citRate.HasValue)
                    {
                        exception.UtprSafeHarbour =
                            new Globe.EtrTypeEtrStatusEtrExceptionUtprSafeHarbour
                            {
                                CitRate = citRate.Value,
                            };
                    }
                    else
                    {
                        errors.Add($"[{fileName}] [{loc}/UTPR] CIT율 형식 오류: '{utprRate}' (0~1 사이 decimal 입력, 예: 0.15)");
                    }
                }

                js.GLoBeTax.Etr.Add(etr);
            }

            // ─── InitialIntActivity (2.3) ─────────────────────────────────
            if (hasInit && TryParseDate(initStartRaw, out var startDate))
            {
                var init = new Globe.InitialIntActivityType { StartDate = startDate };

                if (!string.IsNullOrEmpty(initRefJur))
                {
                    if (TryParseEnum<Globe.CountryCodeType>(initRefJur, out var refCode))
                    {
                        init.ReferenceJurisdiction =
                            new Globe.InitialIntActivityTypeReferenceJurisdiction
                            {
                                ResCountryCode = refCode,
                                TangibleAssetValue = initRefAsset ?? "0",
                            };
                    }
                    else
                        errors.Add(
                            $"[{fileName}] [{loc}/2.3] 준거국가 코드 파싱 실패: '{initRefJur}'"
                        );
                }

                if (!string.IsNullOrEmpty(initNumJur))
                    init.RfyNumberOfJurisdictions = initNumJur;

                // ─── item 5: 준거국가 외 국가 유형자산 (통합 셀 M46) ────────
                // 포맷: "국가코드(ISO2),유형자산값" × N, 국가 간 ';' 구분
                // 예: "KR, 100000000; US, 50000000"
                if (!string.IsNullOrEmpty(initOtherJurRaw))
                    ParseOtherJurisdictionsCell(
                        initOtherJurRaw,
                        init,
                        b2 + 21,
                        errors,
                        fileName,
                        loc
                    );

                js.GLoBeTax.InitialIntActivity = init;
            }
        }

        /// <summary>
        /// 준거국가 외 국가 유형자산 통합 셀 파싱.
        /// "국가,값; 국가,값" → OtherJurisdiction[]
        /// </summary>
        private void ParseOtherJurisdictionsCell(
            string cellValue,
            Globe.InitialIntActivityType init,
            int row,
            List<string> errors,
            string fileName,
            string loc
        )
        {
            var entries = cellValue.Split(
                ';',
                StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
            );

            for (int i = 0; i < entries.Length; i++)
            {
                var parts = entries[i].Split(',', StringSplitOptions.TrimEntries);
                if (parts.Length == 0 || string.IsNullOrEmpty(parts[0]))
                    continue;

                var countryRaw = parts[0];
                var assetValue = parts.Length >= 2 ? parts[1] : "";

                if (!TryParseEnum<Globe.CountryCodeType>(countryRaw, out var otherCode))
                {
                    errors.Add(
                        $"[{fileName}] [{loc}/2.3/외국{i + 1}] 국가코드 파싱 실패: '{countryRaw}' (M{row})"
                    );
                    continue;
                }

                var other = new Globe.InitialIntActivityTypeOtherJurisdiction
                {
                    TangibleAssetValue = assetValue,
                };
                other.ResCountryCode.Add(otherCode);
                init.OtherJurisdiction.Add(other);
            }
        }

        /// <summary>
        /// "국가코드[, (하위그룹유형, TIN, TIN유형, 발급국가)]" 형식 파싱.
        /// 예: "KR, (GIR1101, 123456790, GIR3001, KR)" 또는 "KR"
        /// </summary>
        private Globe.JurisdictionSectionTypeJurWithTaxingRights ParseJwrEntry(
            string entry,
            List<string> errors,
            string fileName,
            string loc
        )
        {
            string countryPart;
            string subgroupPart = null;

            // 쉼표는 있지만 괄호가 없으면 → 잘못된 형식 (복수 항목은 세미콜론으로 구분해야 함)
            var parenIdx = entry.IndexOf('(');
            if (parenIdx < 0 && entry.Contains(','))
            {
                errors.Add(
                    $"[{fileName}] [{loc}] 과세권 국가 형식 오류: '{entry}' — 복수 항목은 세미콜론(;)으로 구분하세요. 예) KR; JP"
                );
                return null;
            }

            if (parenIdx >= 0)
            {
                countryPart = entry[..parenIdx].Trim().TrimEnd(',').Trim();
                var closeIdx = entry.IndexOf(')');
                if (closeIdx > parenIdx)
                    subgroupPart = entry[(parenIdx + 1)..closeIdx].Trim();
            }
            else
            {
                countryPart = entry.Trim();
            }

            if (!TryParseEnum<Globe.CountryCodeType>(countryPart, out var countryCode))
            {
                errors.Add($"[{fileName}] [{loc}] 과세권 국가 코드 '{countryPart}' 파싱 실패");
                return null;
            }

            var jwr = new Globe.JurisdictionSectionTypeJurWithTaxingRights
            {
                JurisdictionName = countryCode,
            };

            if (!string.IsNullOrEmpty(subgroupPart))
            {
                var parts = subgroupPart.Split(',', StringSplitOptions.TrimEntries);
                var subgroup = new Globe.JurisdictionSectionTypeJurWithTaxingRightsSubgroup();

                if (parts.Length >= 1 && !string.IsNullOrEmpty(parts[0]))
                    SetEnum<Globe.TypeofSubGroupEnumType>(
                        parts[0],
                        v => subgroup.TypeofSubGroup.Add(v),
                        errors,
                        fileName,
                        new MappingEntry { Label = $"[{loc}] 과세권/하위그룹유형" }
                    );

                var tinVal = parts.Length >= 2 ? parts[1] : null;
                var tinTypeStr = parts.Length >= 3 ? parts[2] : null;
                var issuedBy = parts.Length >= 4 ? parts[3] : null;

                if (!string.IsNullOrEmpty(tinVal))
                {
                    var tin = new Globe.TinType { Value = tinVal };
                    if (
                        !string.IsNullOrEmpty(tinTypeStr)
                        && TryParseEnum<Globe.TinEnumType>(tinTypeStr, out var tinEnum)
                    )
                    {
                        tin.TypeOfTin = tinEnum;
                        tin.TypeOfTinSpecified = true;
                    }
                    if (
                        !string.IsNullOrEmpty(issuedBy)
                        && TryParseEnum<Globe.CountryCodeType>(issuedBy, out var issuedByCode)
                    )
                    {
                        tin.IssuedBy = issuedByCode;
                        tin.IssuedBySpecified = true;
                    }
                    subgroup.Tin = tin;
                }
                else
                {
                    subgroup.Tin = NoTin();
                }

                jwr.Subgroup.Add(subgroup);
            }

            return jwr;
        }

        private static bool RowContains(IXLWorksheet ws, int row, string text)
        {
            for (int col = 2; col <= 6; col++)
            {
                var val = ws.Cell(row, col).GetString();
                if (val?.Contains(text) == true)
                    return true;
            }
            return false;
        }

        private static void TryAddFinancialData(
            Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalc dmCalc,
            DateTime year,
            string revenue,
            string globeRevenue,
            string netGlobeIncome,
            string fanil
        )
        {
            if (
                string.IsNullOrEmpty(globeRevenue)
                && string.IsNullOrEmpty(revenue)
                && string.IsNullOrEmpty(netGlobeIncome)
                && string.IsNullOrEmpty(fanil)
            )
                return;

            dmCalc.FinancialData.Add(
                new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcFinancialData
                {
                    Year = year,
                    Revenue = NullIfEmpty(revenue),
                    GlobeRevenue = NullIfEmpty(globeRevenue),
                    NetGlobeIncome = NullIfEmpty(netGlobeIncome),
                    Fanil = NullIfEmpty(fanil),
                }
            );
        }

        private static void TryAddSimpleFinancialData(
            Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalc dmCalc,
            DateTime year,
            string revenue,
            string simplifiedTax
        )
        {
            if (string.IsNullOrEmpty(revenue) && string.IsNullOrEmpty(simplifiedTax))
                return;

            // 간소화 세이프하버 데이터는 CbCR 기반이라 GlobeRevenue/FANIL 별도값 없음.
            // XSD에서 GlobeRevenue/NetGlobeIncome/FANIL 모두 required이므로
            // 간소화 입력값을 그대로 복제 (Revenue=GlobeRevenue, NetGlobeIncome=FANIL).
            dmCalc.FinancialData.Add(
                new Globe.EtrTypeEtrStatusEtrExceptionDeminimisSimplifiedNmceCalcFinancialData
                {
                    Year = year,
                    Revenue = NullIfEmpty(revenue),
                    GlobeRevenue = NullIfEmpty(revenue),
                    NetGlobeIncome = NullIfEmpty(simplifiedTax),
                    Fanil = NullIfEmpty(simplifiedTax),
                }
            );
        }
    }
}
