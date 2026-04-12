using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// entity_template.xlsx "구성기업 계산" 시트 → CEComputation 매핑.
    /// 현재 구현: CE TIN 식별 + 3.2.4(a) SimplCalculations + 3.2.4(b) AggregatedReporting.
    /// </summary>
    public class Mapping_EntityCe : MappingBase
    {
        public Mapping_EntityCe() : base(null) { }

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        ) => MapWithJur(ws, globe, errors, fileName, null);

        /// <summary>
        /// jurCode가 있으면 해당 JurisdictionSection에 CEComputation 추가.
        /// null이면 CE TIN issuedBy로 폴백.
        /// </summary>
        public void MapWithJur(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName,
            Globe.CountryCodeType? jurCode,
            string subGroupTin = null
        )
        {
            // ── CE TIN 식별 (3.2.4.1(a) 첫 번째 납세자번호 항목) ─────────
            var rTin = FindRow(ws, "1. 구성기업 또는 공동기업그룹 기업의 납세자번호");
            if (rTin < 0)
            {
                errors.Add($"[{fileName}] CE TIN 항목 없음 — '구성기업 계산' 시트에서 납세자번호 행을 찾을 수 없음");
                return;
            }

            var ceTinRaw = ws.Cell(rTin, 13).GetString()?.Trim(); // M열
            if (string.IsNullOrEmpty(ceTinRaw))
            {
                errors.Add($"[{fileName}] CE TIN 미입력 — '구성기업 계산' 3.2.4.1(a) 납세자번호 M열에 '값,GIR3001,KR' 형식으로 입력 필요");
                return;
            }

            var ceTin = ParseTin(ceTinRaw);

            // ── JurisdictionSection 조회: 폴더 기준 국가코드 우선, 없으면 CE TIN issuedBy 폴백 ──
            Globe.CountryCodeType lookupCode;
            if (jurCode.HasValue)
            {
                lookupCode = jurCode.Value;
            }
            else if (ceTin.IssuedBySpecified)
            {
                lookupCode = ceTin.IssuedBy;
            }
            else
            {
                errors.Add($"[{fileName}] CE TIN '{ceTinRaw}' 형식 오류 — '납세자번호값,GIR3001,KR' 처럼 발급국가(ISO 2자리)까지 입력 필요");
                return;
            }

            var js = globe.GlobeBody.JurisdictionSection
                .FirstOrDefault(s => s.Jurisdiction == lookupCode);
            if (js == null)
            {
                errors.Add($"[{fileName}] JurisdictionSection({lookupCode}) 없음 — 합산단위 파일과 같은 폴더에 위치하는지 확인");
                return;
            }

            // ── ETR 찾기: SubGroup TIN 일치 우선, 없으면 첫 번째 ─────────
            Globe.EtrType etr;
            if (!string.IsNullOrEmpty(subGroupTin))
                etr = js.GLoBeTax.Etr.FirstOrDefault(e => e.SubGroup?.Tin?.Value == subGroupTin)
                   ?? js.GLoBeTax.Etr.FirstOrDefault();
            else
                etr = js.GLoBeTax.Etr.FirstOrDefault();

            if (etr == null)
            {
                etr = new Globe.EtrType { EtrStatus = new Globe.EtrTypeEtrStatus() };
                js.GLoBeTax.Etr.Add(etr);
            }
            etr.EtrStatus ??= new Globe.EtrTypeEtrStatus();
            etr.EtrStatus.EtrComputation ??= new Globe.EtrComputationType();

            // ── CEComputation 찾기 또는 생성 ─────────────────────────────
            var ceComp = etr.EtrStatus.EtrComputation.CeComputation
                .FirstOrDefault(c => c.Tin?.Value == ceTin.Value);
            if (ceComp == null)
            {
                ceComp = new Globe.EtrComputationTypeCeComputation { Tin = ceTin };
                etr.EtrStatus.EtrComputation.CeComputation.Add(ceComp);
            }

            // ── 3.2.4(a) 전환기 국가별 간소화 신고체계 선택 ──────────────
            Map324a(ws, ceComp, errors, fileName);

            // ── 3.2.4(b) 연결납세그룹 통합신고 ───────────────────────────
            Map324b(ws, ceComp, errors, fileName);

            // ── 3.2.4.1(a) 배분 후 회계상 순손익(FANIL) + 글로벌최저한세소득결손(Total) ──
            Map3241a(ws, ceComp, errors, fileName);

            // ── 3.2.4.1(b) 본점·고정사업장 간 손익 배분 → MainEntityPEandFte ──
            Map3241b(ws, ceComp, errors, fileName);

            // ── 3.2.4.1(c) 국가간 손익 조정 → CrossBorderAdjustments ────
            Map3241c(ws, ceComp, errors, fileName);

            // ── 3.2.4.1(d) 최종모기업 글로벌최저한세소득 감액 → UpeAdjustments ──
            Map3241d(ws, ceComp, errors, fileName);

            // ── 3.2.4.2(a) 당기법인세비용 + 조정항목(a~q) + 조정대상조세 ──
            Map3242a(ws, ceComp, errors, fileName);

            // ── 3.2.4.2(b) 대상조세 국가간 배분 → IncomeTax + CrossAllocation[] ──
            Map3242b(ws, ceComp, errors, fileName);

            // ── 3.2.4.2(c) 이연법인세 → DeferTaxAdjustAmt ───────────────────
            Map3242c(ws, ceComp, errors, fileName);

            // ── 3.2.4.3 구성기업별 선택 ──────────────────────────────────────
            Map3243(ws, ceComp, errors, fileName);

            // ── 3.2.4.4 국제해운소득·결손 제외 ──────────────────────────────
            Map3244(ws, ceComp, errors, fileName);

            // ── 3.2.4.5 과세분배방법 적용 선택 관련 정보 ─────────────────────
            Map3245(ws, ceComp, errors, fileName);

            // ── 3.2.4.6 그 밖의 회계기준 ─────────────────────────────────────
            Map3246(ws, ceComp, errors, fileName);

            // ── 3.4.1/3.4.2 "추가세액" 별도 시트 ────────────────────────────
            if (ws.Workbook.TryGetWorksheet("추가세액", out var topUpWs))
            {
                Map341(topUpWs, js, errors, fileName);
                Map342(topUpWs, js, errors, fileName);
            }
        }

        // ─── 3.2.4(a): row에서 K열 여/부 → Elections.SimplCalculations ──
        private static void Map324a(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var r = FindRow(ws, "1. 귀 다국적기업그룹은");
            if (r < 0) return;
            var v = ws.Cell(r, 15).GetString()?.Trim(); // O열
            if (string.IsNullOrEmpty(v)) return;

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            ceComp.Elections.SimplCalculations = ParseBool(v);
            ceComp.Elections.SimplCalculationsSpecified = true;
        }

        // ─── 3.2.4(b): 헤더 다음 행들에서 B=GroupTIN, K=EntityTINs ──────
        // 헤더행: B="1. 연결납세그룹(납세자번호)", K="2. 해당 연결납세그룹에 포함된 기업..."
        // 데이터행: B=GroupTIN ("값,유형,발급국가"), K=EntityTIN (행마다 1개)
        private static void Map324b(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "1. 연결납세그룹(납세자번호)");
            if (rHdr < 0) return;

            string groupTinRaw = null;
            var entityTins = new System.Collections.Generic.List<Globe.TinType>();

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 10;
            for (int r = rHdr + 1; r <= System.Math.Min(rHdr + 20, lastRow); r++)
            {
                var bVal = ws.Cell(r, 2).GetString()?.Trim();   // B: GroupTIN
                var kVal = ws.Cell(r, 11).GetString()?.Trim();  // K: EntityTIN

                if (string.IsNullOrEmpty(bVal) && string.IsNullOrEmpty(kVal)) break;

                if (!string.IsNullOrEmpty(bVal) && groupTinRaw == null)
                    groupTinRaw = bVal;

                if (!string.IsNullOrEmpty(kVal))
                    entityTins.Add(ParseTin(kVal));
            }

            if (groupTinRaw == null) return;

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            var ar = new Globe.EtrComputationTypeCeComputationElectionsAggregatedReporting
            {
                TaxConsolGroupTin = ParseTin(groupTinRaw)
            };
            foreach (var t in entityTins)
                ar.EntityTin.Add(t);

            ceComp.Elections.AggregatedReporting = ar;
        }

        // ─── 3.2.4.2(a): 당기법인세비용 → AdjustedIncomeTax.Total
        //                 (a)~(q) 17개 조정항목 → AdjustedCoveredTax.Adjustments
        //                 조정대상조세 합계 → AdjustedCoveredTax.Total
        private static readonly Globe.CurrentAdjustedTaxEnumType[] TaxAdjustmentItems =
        {
            Globe.CurrentAdjustedTaxEnumType.Gir2401, Globe.CurrentAdjustedTaxEnumType.Gir2402,
            Globe.CurrentAdjustedTaxEnumType.Gir2403, Globe.CurrentAdjustedTaxEnumType.Gir2404,
            Globe.CurrentAdjustedTaxEnumType.Gir2405, Globe.CurrentAdjustedTaxEnumType.Gir2406,
            Globe.CurrentAdjustedTaxEnumType.Gir2407, Globe.CurrentAdjustedTaxEnumType.Gir2408,
            Globe.CurrentAdjustedTaxEnumType.Gir2409, Globe.CurrentAdjustedTaxEnumType.Gir2410,
            Globe.CurrentAdjustedTaxEnumType.Gir2411, Globe.CurrentAdjustedTaxEnumType.Gir2412,
            Globe.CurrentAdjustedTaxEnumType.Gir2413, Globe.CurrentAdjustedTaxEnumType.Gir2414,
            Globe.CurrentAdjustedTaxEnumType.Gir2415, Globe.CurrentAdjustedTaxEnumType.Gir2416,
            Globe.CurrentAdjustedTaxEnumType.Gir2417,
        };

        private static void Map3242a(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.2 조정대상조세");
            if (rSec < 0) return;

            // 2. 배분 후 당기법인세비용 (O열) → AdjustedIncomeTax.Total
            var rIncomeTax = FindRow(ws, "배분 후 당기법인세비용", rSec);
            if (rIncomeTax >= 0)
            {
                var v = ws.Cell(rIncomeTax, 15).GetString()?.Trim(); // O열
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.AdjustedIncomeTax ??= new Globe.EtrComputationTypeCeComputationAdjustedIncomeTax();
                    ceComp.AdjustedIncomeTax.Total = v;
                }
            }

            // 3. 조정사항 (a)~(q) 17개 → AdjustedCoveredTax.Adjustments (GIR2401~2417)
            // O(15)=가산액, Q(17)=차감액
            var rAdjHdr = FindRow(ws, "3. 조정사항", rSec);
            if (rAdjHdr >= 0)
            {
                ceComp.AdjustedCoveredTax ??= new Globe.EtrComputationTypeCeComputationAdjustedCoveredTax();
                for (int i = 0; i < TaxAdjustmentItems.Length; i++)
                {
                    var adds = ws.Cell(rAdjHdr + 1 + i, 15).GetString()?.Trim(); // O: 가산액
                    var reds = ws.Cell(rAdjHdr + 1 + i, 17).GetString()?.Trim(); // Q: 차감액
                    if (string.IsNullOrEmpty(adds) && string.IsNullOrEmpty(reds)) continue;

                    var adj = new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxAdjustments
                    {
                        AdjustmentItem = TaxAdjustmentItems[i]
                    };
                    if (!string.IsNullOrEmpty(adds)) adj.Amount.Add(adds);
                    if (!string.IsNullOrEmpty(reds)) adj.Amount.Add("-" + reds);
                    ceComp.AdjustedCoveredTax.Adjustments.Add(adj);
                }
            }

            // 4. 조정대상조세 (O열) → AdjustedCoveredTax.Total
            var rTotal = FindRow(ws, "4. 조정대상조세", rSec);
            if (rTotal >= 0)
            {
                var v = ws.Cell(rTotal, 15).GetString()?.Trim(); // O열
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.AdjustedCoveredTax ??= new Globe.EtrComputationTypeCeComputationAdjustedCoveredTax();
                    ceComp.AdjustedCoveredTax.Total = v;
                }
            }
        }

        // ─── 3.2.4.2(b): 대상조세 국가간 배분 → AdjustedIncomeTax.IncomeTax + CrossAllocation[] ──
        // 헤더행: "대상조세 국가간 배분", 컬럼헤더 +1행, 데이터 +2행~ (행추가 가능)
        // C(3)=배분전(IncomeTax, 첫 행만), E(5)=Basis, G(7)=OtherTIN, J(10)=ResCountry
        // L(12)=Additions, N(14)=Reductions
        private static void Map3242b(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "대상조세 국가간 배분");
            if (rHdr < 0) return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            bool firstRow = true;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 30, lastRow); r++)
            {
                var basisRaw  = ws.Cell(r, 5).GetString()?.Trim();   // E: 조정 근거
                var otherTinRaw = ws.Cell(r, 7).GetString()?.Trim(); // G: 상대TIN
                var adds      = ws.Cell(r, 12).GetString()?.Trim();  // L: 가산액
                var reds      = ws.Cell(r, 14).GetString()?.Trim();  // N: 차감액

                if (string.IsNullOrEmpty(basisRaw) && string.IsNullOrEmpty(otherTinRaw)
                    && string.IsNullOrEmpty(adds) && string.IsNullOrEmpty(reds)) break;

                // 첫 데이터행 C(3): 배분 전 대상조세 → AdjustedIncomeTax.IncomeTax
                if (firstRow)
                {
                    firstRow = false;
                    var preTax = ws.Cell(r, 3).GetString()?.Trim();
                    if (!string.IsNullOrEmpty(preTax))
                    {
                        ceComp.AdjustedIncomeTax ??= new Globe.EtrComputationTypeCeComputationAdjustedIncomeTax();
                        ceComp.AdjustedIncomeTax.IncomeTax = preTax;
                    }
                }

                if (string.IsNullOrEmpty(basisRaw)) continue;

                if (!TryParseEnum<Globe.AdjustedBasisEnumType>(basisRaw, out var basis))
                {
                    errors.Add($"[{fileName}] 3.2.4.2(b) 조정 근거 '{basisRaw}' 파싱 실패 — GIR2201~GIR2205 중 하나여야 함");
                    continue;
                }

                var item = new Globe.EtrComputationTypeCeComputationAdjustedIncomeTaxCrossAllocation();
                item.Basis.Add(basis);
                if (!string.IsNullOrEmpty(otherTinRaw)) item.OtherTin = ParseTin(otherTinRaw);
                if (!string.IsNullOrEmpty(adds)) item.Additions = adds;
                if (!string.IsNullOrEmpty(reds)) item.Reductions = reds;

                var ccRaw = ws.Cell(r, 10).GetString()?.Trim(); // J: 소재지국
                if (!string.IsNullOrEmpty(ccRaw)
                    && TryParseEnum<Globe.CountryCodeType>(ccRaw, out var cc))
                {
                    item.ResCountryCode = cc;
                }

                ceComp.AdjustedIncomeTax ??= new Globe.EtrComputationTypeCeComputationAdjustedIncomeTax();
                ceComp.AdjustedIncomeTax.CrossAllocation.Add(item);
            }
        }

        // ─── 3.2.4.2(c): 이연법인세 → DeferTaxAdjustAmt ──────────────────
        // 헤더: "이연법인세비용 금액(글로벌최저한세 기준)" → DeferTaxExpense
        // 조정항목 (a)~(p) 16개: "3. 이연법인세비용의 조정" +1~ → Adjustment[]
        //   O(15)=가산액, Q(17)=차감액(양수 입력, "-" 접두)
        // 낮은 세율 재계산 → GIR2511 항목의 Recast.Lower
        // 높은 세율 재계산 → GIR2512 항목의 Recast.Higher
        // 총이연법인세조정금액 → Total
        private static readonly Globe.DeferredAdjustedTaxEnumType[] DeferredTaxAdjItems =
        {
            Globe.DeferredAdjustedTaxEnumType.Gir2501, Globe.DeferredAdjustedTaxEnumType.Gir2502,
            Globe.DeferredAdjustedTaxEnumType.Gir2503, Globe.DeferredAdjustedTaxEnumType.Gir2504,
            Globe.DeferredAdjustedTaxEnumType.Gir2505, Globe.DeferredAdjustedTaxEnumType.Gir2506,
            Globe.DeferredAdjustedTaxEnumType.Gir2507, Globe.DeferredAdjustedTaxEnumType.Gir2508,
            Globe.DeferredAdjustedTaxEnumType.Gir2509, Globe.DeferredAdjustedTaxEnumType.Gir2510,
            Globe.DeferredAdjustedTaxEnumType.Gir2511, Globe.DeferredAdjustedTaxEnumType.Gir2512,
            Globe.DeferredAdjustedTaxEnumType.Gir2513, Globe.DeferredAdjustedTaxEnumType.Gir2514,
            Globe.DeferredAdjustedTaxEnumType.Gir2515, Globe.DeferredAdjustedTaxEnumType.Gir2516,
        };

        private static void Map3242c(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "(c) 이연법인세");
            if (rSec < 0) return;

            // 2. 이연법인세비용 금액(글로벌최저한세 기준) O열 → DeferTaxExpense
            var rExpense = FindRow(ws, "이연법인세비용 금액", rSec);
            string deferExpenseVal = null;
            if (rExpense >= 0)
                deferExpenseVal = ws.Cell(rExpense, 15).GetString()?.Trim();

            // 3. 조정항목 (a)~(p) 16개: O(15)=가산액, Q(17)=차감액
            var rAdjHdr = FindRow(ws, "이연법인세비용의 조정", rSec);
            var adjItems = new System.Collections.Generic.Dictionary<Globe.DeferredAdjustedTaxEnumType,
                Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment>();

            if (rAdjHdr >= 0)
            {
                for (int i = 0; i < DeferredTaxAdjItems.Length; i++)
                {
                    var adds = ws.Cell(rAdjHdr + 1 + i, 15).GetString()?.Trim(); // O: 가산액
                    var reds = ws.Cell(rAdjHdr + 1 + i, 17).GetString()?.Trim(); // Q: 차감액
                    if (string.IsNullOrEmpty(adds) && string.IsNullOrEmpty(reds)) continue;

                    var adj = new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment
                    {
                        AdjustmentItem = DeferredTaxAdjItems[i]
                    };
                    if (!string.IsNullOrEmpty(adds)) adj.Amount.Add(adds);
                    if (!string.IsNullOrEmpty(reds)) adj.Amount.Add("-" + reds);
                    adjItems[DeferredTaxAdjItems[i]] = adj;
                }
            }

            // 4. 낮은 세율 재계산 차이금액 → GIR2511 항목의 Recast.Lower
            var rRecastDown = FindRow(ws, "최저한세율보다 낮은 세율", rSec);
            if (rRecastDown >= 0)
            {
                var v = ws.Cell(rRecastDown, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(v))
                {
                    if (!adjItems.TryGetValue(Globe.DeferredAdjustedTaxEnumType.Gir2511, out var adj2511))
                    {
                        adj2511 = new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment
                        {
                            AdjustmentItem = Globe.DeferredAdjustedTaxEnumType.Gir2511
                        };
                        adjItems[Globe.DeferredAdjustedTaxEnumType.Gir2511] = adj2511;
                    }
                    adj2511.Recast ??= new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentRecast();
                    adj2511.Recast.Lower = v;
                }
            }

            // 5. 높은 세율 재계산 차이금액 → GIR2512 항목의 Recast.Higher
            var rRecastUp = FindRow(ws, "최저한세율보다 높은 세율", rSec);
            if (rRecastUp >= 0)
            {
                var v = ws.Cell(rRecastUp, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(v))
                {
                    if (!adjItems.TryGetValue(Globe.DeferredAdjustedTaxEnumType.Gir2512, out var adj2512))
                    {
                        adj2512 = new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment
                        {
                            AdjustmentItem = Globe.DeferredAdjustedTaxEnumType.Gir2512
                        };
                        adjItems[Globe.DeferredAdjustedTaxEnumType.Gir2512] = adj2512;
                    }
                    adj2512.Recast ??= new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentRecast();
                    adj2512.Recast.Higher = v;
                }
            }

            // 6. 총이연법인세조정금액 → Total
            var rTotal = FindRow(ws, "총이연법인세조정금액", rSec);
            string totalVal = null;
            if (rTotal >= 0)
                totalVal = ws.Cell(rTotal, 15).GetString()?.Trim();

            // 값이 하나라도 있으면 DeferTaxAdjustAmt 생성
            if (!string.IsNullOrEmpty(deferExpenseVal) || adjItems.Count > 0 || !string.IsNullOrEmpty(totalVal))
            {
                ceComp.AdjustedCoveredTax ??= new Globe.EtrComputationTypeCeComputationAdjustedCoveredTax();
                var dta = ceComp.AdjustedCoveredTax.DeferTaxAdjustAmt
                       ?? new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmt();
                ceComp.AdjustedCoveredTax.DeferTaxAdjustAmt = dta;

                if (!string.IsNullOrEmpty(deferExpenseVal)) dta.DeferTaxExpense = deferExpenseVal;
                if (!string.IsNullOrEmpty(totalVal)) dta.Total = totalVal;

                // Adjustment[]는 enum 순서(GIR2501~2516)대로 추가
                foreach (var item in DeferredTaxAdjItems)
                    if (adjItems.TryGetValue(item, out var a)) dta.Adjustment.Add(a);
            }
        }

        // ─── 3.2.4.1(a): FANIL → AdjustedFanil.Fanil
        //                  (a)~(z) 26개 조정항목 → NetGlobeIncome.Adjustments[]
        //                  글로벌최저한세소득결손 합계 → NetGlobeIncome.Total
        private static readonly Globe.AdjustmentItemEnumType[] AdjustmentItems =
        {
            Globe.AdjustmentItemEnumType.Gir2001, Globe.AdjustmentItemEnumType.Gir2002,
            Globe.AdjustmentItemEnumType.Gir2003, Globe.AdjustmentItemEnumType.Gir2004,
            Globe.AdjustmentItemEnumType.Gir2005, Globe.AdjustmentItemEnumType.Gir2006,
            Globe.AdjustmentItemEnumType.Gir2007, Globe.AdjustmentItemEnumType.Gir2008,
            Globe.AdjustmentItemEnumType.Gir2009, Globe.AdjustmentItemEnumType.Gir2010,
            Globe.AdjustmentItemEnumType.Gir2011, Globe.AdjustmentItemEnumType.Gir2012,
            Globe.AdjustmentItemEnumType.Gir2013, Globe.AdjustmentItemEnumType.Gir2014,
            Globe.AdjustmentItemEnumType.Gir2015, Globe.AdjustmentItemEnumType.Gir2016,
            Globe.AdjustmentItemEnumType.Gir2017, Globe.AdjustmentItemEnumType.Gir2018,
            Globe.AdjustmentItemEnumType.Gir2019, Globe.AdjustmentItemEnumType.Gir2020,
            Globe.AdjustmentItemEnumType.Gir2021, Globe.AdjustmentItemEnumType.Gir2022,
            Globe.AdjustmentItemEnumType.Gir2023, Globe.AdjustmentItemEnumType.Gir2024,
            Globe.AdjustmentItemEnumType.Gir2025, Globe.AdjustmentItemEnumType.Gir2026,
        };

        private static void Map3241a(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            // M16 "배분 후 회계상 순손익" → AdjustedFanil.Total
            var rFanil = FindRow(ws, "배분 후 회계상 순손익");
            if (rFanil >= 0)
            {
                var v = ws.Cell(rFanil, 13).GetString()?.Trim();
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                    ceComp.AdjustedFanil.Total = v;
                }
            }

            // (a)~(z) 26개 조정항목 (rows 18~43) → NetGlobeIncome.Adjustments[]
            // 헤더행 "3. 조정사항" 바로 다음 행부터 26행이 고정 구조
            var rAdjHdr = FindRow(ws, "3. 조정사항");
            if (rAdjHdr >= 0)
            {
                ceComp.NetGlobeIncome ??= new Globe.EtrComputationTypeCeComputationNetGlobeIncome();
                for (int i = 0; i < AdjustmentItems.Length; i++)
                {
                    var adds = ws.Cell(rAdjHdr + 1 + i, 13).GetString()?.Trim(); // M: 가산액
                    var reds = ws.Cell(rAdjHdr + 1 + i, 16).GetString()?.Trim(); // P: 차감액
                    if (string.IsNullOrEmpty(adds) && string.IsNullOrEmpty(reds)) continue;

                    var adj = new Globe.EtrComputationTypeCeComputationNetGlobeIncomeAdjustments
                    {
                        AdjustmentItem = AdjustmentItems[i]
                    };
                    if (!string.IsNullOrEmpty(adds)) adj.Amount.Add(adds);
                    if (!string.IsNullOrEmpty(reds)) adj.Amount.Add("-" + reds);
                    ceComp.NetGlobeIncome.Adjustments.Add(adj);
                }
            }

            // 글로벌최저한세소득결손 합계 (row 44) → NetGlobeIncome.Total
            var rTotal = FindRow(ws, "4. 구성기업 또는 공동기업그룹 기업의 글로벌최저한세소득");
            if (rTotal >= 0)
            {
                var v = ws.Cell(rTotal, 13).GetString()?.Trim();
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.NetGlobeIncome ??= new Globe.EtrComputationTypeCeComputationNetGlobeIncome();
                    ceComp.NetGlobeIncome.Total = v;
                }
            }
        }

        // ─── 3.2.4.1(b): 본점·고정사업장 간 배분 → MainEntityPEandFte ──
        // 헤더행: "(b) 본점과 고정사업장", 컬럼헤더 +1행, 데이터 +2행~ (행추가 가능)
        // E(5)=Basis(GIR1701~1704), G(7)=OtherTIN, J(10)=ResCountry, L(12)=Additions, N(14)=Reductions
        // 첫 데이터행 C(3)="2. 조정 전 회계상 순손익" 값 → AdjustedFanil.Fanil
        // (b) 섹션 없거나 C(3) 비어있으면 M16 "배분 후 회계상 순손익" 으로 폴백
        private static void Map3241b(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "본점과 고정사업장");

            // AdjustedFanil.Fanil: section (b) 첫 데이터행 C(3), 없으면 M16 폴백
            string fanilVal = null;
            if (rHdr >= 0)
                fanilVal = ws.Cell(rHdr + 2, 3).GetString()?.Trim(); // C열: 조정 전 회계상 순손익
            if (string.IsNullOrEmpty(fanilVal))
            {
                var rM16 = FindRow(ws, "배분 후 회계상 순손익");
                if (rM16 >= 0)
                    fanilVal = ws.Cell(rM16, 13).GetString()?.Trim();
            }
            if (!string.IsNullOrEmpty(fanilVal))
            {
                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Fanil = fanilVal;
            }

            if (rHdr < 0) return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 15, lastRow); r++)
            {
                var basisRaw = ws.Cell(r, 5).GetString()?.Trim();  // E: 조정근거
                var otherTinRaw = ws.Cell(r, 7).GetString()?.Trim(); // G: 상대TIN
                var adds = ws.Cell(r, 12).GetString()?.Trim();       // L: 가산액
                var reds = ws.Cell(r, 14).GetString()?.Trim();       // N: 차감액

                if (string.IsNullOrEmpty(basisRaw) && string.IsNullOrEmpty(otherTinRaw)
                    && string.IsNullOrEmpty(adds) && string.IsNullOrEmpty(reds)) break;
                if (string.IsNullOrEmpty(basisRaw)) continue;

                if (!TryParseEnum<Globe.MainEntityPEandFteBasisEnumType>(basisRaw, out var basis))
                {
                    errors.Add($"[{fileName}] 3.2.4.1(b) 조정근거 '{basisRaw}' 파싱 실패 — GIR1701~GIR1704 중 하나여야 함");
                    continue;
                }

                var item = new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentMainEntityPEandFte
                {
                    Basis = basis,
                    OtherTin = ParseTin(otherTinRaw ?? ""),
                    Additions = adds ?? "0",
                    Reductions = reds ?? "0"
                };

                var ccRaw = ws.Cell(r, 10).GetString()?.Trim(); // J: 소재지국
                if (!string.IsNullOrEmpty(ccRaw)
                    && TryParseEnum<Globe.CountryCodeType>(ccRaw, out var cc))
                {
                    item.ResCountryCode = cc;
                    item.ResCountryCodeSpecified = true;
                }

                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Adjustment ??= new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustment();
                ceComp.AdjustedFanil.Adjustment.MainEntityPEandFte.Add(item);
            }
        }

        // ─── 3.2.4.1(c): 국가간 손익 조정 → CrossBorderAdjustments ─────
        // 헤더행: "(c) 국가간 손익 조정", 컬럼헤더 +1행, 데이터 +2행~
        // C(3)=Basis, F(6)=OtherTIN, I(9)=ResCountryCode, L(12)=Additions, P(16)=Reductions
        private static void Map3241c(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "국가간 손익 조정");
            if (rHdr < 0) return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 15, lastRow); r++)
            {
                var basisRaw = ws.Cell(r, 3).GetString()?.Trim();   // C: 조정근거
                var otherTinRaw = ws.Cell(r, 6).GetString()?.Trim(); // F: 상대TIN
                var adds = ws.Cell(r, 12).GetString()?.Trim();       // L: 가산액
                var reds = ws.Cell(r, 16).GetString()?.Trim();       // P: 차감액

                if (string.IsNullOrEmpty(basisRaw) && string.IsNullOrEmpty(otherTinRaw)
                    && string.IsNullOrEmpty(adds) && string.IsNullOrEmpty(reds)) break;
                if (string.IsNullOrEmpty(basisRaw)) continue;

                if (!TryParseEnum<Globe.CrossBorderAdjustmentsEnumType>(basisRaw, out var basis))
                {
                    errors.Add($"[{fileName}] 3.2.4.1(c) 조정근거 '{basisRaw}' 파싱 실패 — GIR1801~GIR1802 중 하나여야 함");
                    continue;
                }

                var item = new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentCrossBorderAdjustments
                {
                    Basis = basis,
                    OtherTin = ParseTin(otherTinRaw ?? ""),
                    Additions = adds,
                    Reductions = reds
                };

                var ccRaw = ws.Cell(r, 9).GetString()?.Trim(); // I: 소재지국
                if (!string.IsNullOrEmpty(ccRaw)
                    && TryParseEnum<Globe.CountryCodeType>(ccRaw, out var cc))
                {
                    item.ResCountryCode = cc;
                    item.ResCountryCodeSpecified = true;
                }

                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Adjustment ??= new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustment();
                ceComp.AdjustedFanil.Adjustment.CrossBorderAdjustments.Add(item);
            }
        }

        // ─── 3.2.4.1(d): 최종모기업 글로벌최저한세소득 감액 → UpeAdjustments ──
        // 헤더행: "최종모기업의 글로벌최저한세소득 감액", 컬럼헤더 +1행, 데이터 +2행~ (행추가 가능)
        // E(5)=Basis, H(8)=소유지분보유자TIN, K(11)=직접보유비율(%), N(14)=소득차감액
        private static void Map3241d(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "최종모기업의 글로벌최저한세소득 감액");
            if (rHdr < 0) return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 20, lastRow); r++)
            {
                var basisRaw  = ws.Cell(r, 5).GetString()?.Trim();   // E: 감액근거
                var ownerRaw  = ws.Cell(r, 8).GetString()?.Trim();   // H: 소유지분보유자TIN
                var pctRaw    = ws.Cell(r, 11).GetString()?.Trim();  // K: 직접보유비율(%)
                var reductRaw = ws.Cell(r, 14).GetString()?.Trim();  // N: 소득차감액

                if (string.IsNullOrEmpty(basisRaw) && string.IsNullOrEmpty(ownerRaw)
                    && string.IsNullOrEmpty(pctRaw) && string.IsNullOrEmpty(reductRaw)) break;
                if (string.IsNullOrEmpty(basisRaw)) continue;

                if (!TryParseEnum<Globe.UpeAdjustmentsBasisEnumType>(basisRaw, out var basis))
                {
                    errors.Add($"[{fileName}] 3.2.4.1(d) 감액근거 '{basisRaw}' 파싱 실패 — GIR1901 등 유효한 코드여야 함");
                    continue;
                }

                var upe = new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustments
                {
                    Basis = basis
                };

                // N열: 소득차감액 → Reductions.Amount
                if (!string.IsNullOrEmpty(reductRaw))
                    upe.Reductions = new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsReductions
                    {
                        Amount = reductRaw
                    };

                // H열 + K열: 소유지분보유자 → IdentificationOfOwners
                if (!string.IsNullOrEmpty(ownerRaw) || !string.IsNullOrEmpty(pctRaw))
                {
                    var owner = ParseOwner(ownerRaw, pctRaw, errors, fileName);
                    if (owner != null)
                        upe.IdentificationOfOwners.Add(owner);
                }

                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Adjustment ??= new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustment();
                ceComp.AdjustedFanil.Adjustment.UpeAdjustments.Add(upe);
            }
        }

        // ─── 소유지분보유자 파싱 ────────────────────────────────────────────
        // 개인: 개인,주주수,거주지국[,세율]
        // 법인: 법인,TIN값,유형,발급국가[,세율[,ExTypeOfEntity코드]]
        //       세율 생략 시 빈 칸 허용: 법인,TIN,유형,발급국가,,ExType
        private static Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwners ParseOwner(
            string ownerRaw, string pctRaw, List<string> errors, string fileName)
        {
            var owner = new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwners();

            // K열: 직접보유비율(%)
            if (!string.IsNullOrEmpty(pctRaw))
            {
                var pctClean = pctRaw.TrimEnd('%').Trim();
                if (decimal.TryParse(pctClean, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var pct))
                    owner.OwnershipPercentage = pct > 1m ? pct / 100m : pct;
            }

            if (string.IsNullOrEmpty(ownerRaw)) return owner;

            var parts = ownerRaw.Split(',');
            var kind = parts[0].Trim();

            if (kind == "개인")
            {
                // 개인,주주수,거주지국[,세율]
                var ind = new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwnersIndOwners
                {
                    NumOfOwners = parts.Length > 1 ? parts[1].Trim() : "1"
                };
                if (parts.Length > 2 && !string.IsNullOrEmpty(parts[2].Trim())
                    && TryParseEnum<Globe.CountryCodeType>(parts[2].Trim(), out var cc))
                {
                    ind.ResCountryCode = cc;
                    ind.ResCountryCodeSpecified = true;
                }
                if (parts.Length > 3 && !string.IsNullOrEmpty(parts[3].Trim())
                    && decimal.TryParse(parts[3].Trim(), System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var rate))
                {
                    ind.TaxRate = rate > 1m ? rate / 100m : rate;
                    ind.TaxRateSpecified = true;
                }
                owner.IndOwners = ind;
            }
            else if (kind == "법인")
            {
                // 법인,TIN값,유형,발급국가[,세율[,ExTypeOfEntity코드]]
                if (parts.Length < 4)
                {
                    errors.Add($"[{fileName}] 3.2.4.1(d) 법인 소유자 형식 오류 '{ownerRaw}' — 법인,TIN값,유형,발급국가[,세율[,ExType]] 형식이어야 함");
                    return owner;
                }
                var tinRaw = $"{parts[1].Trim()},{parts[2].Trim()},{parts[3].Trim()}";
                var tin = ParseTin(tinRaw);
                var entity = new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwnersEntityOwner
                {
                    Tin = tin,
                    ResCountryCode = tin.IssuedBySpecified ? tin.IssuedBy : default
                };
                // parts[4]: 세율 (빈 문자열이면 스킵)
                if (parts.Length > 4 && !string.IsNullOrEmpty(parts[4].Trim())
                    && decimal.TryParse(parts[4].Trim(), System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var rate))
                {
                    entity.TaxRate = rate > 1m ? rate / 100m : rate;
                    entity.TaxRateSpecified = true;
                }
                // parts[5]: ExTypeOfEntity (세율 스킵해도 파싱)
                if (parts.Length > 5 && !string.IsNullOrEmpty(parts[5].Trim())
                    && TryParseEnum<Globe.ExTypeOfEntityEnumType>(parts[5].Trim(), out var exType))
                {
                    entity.ExTypeOfEntity = exType;
                    entity.ExTypeOfEntitySpecified = true;
                }
                owner.EntityOwner = entity;
            }
            else
            {
                errors.Add($"[{fileName}] 3.2.4.1(d) 소유자 유형 '{kind}' 오류 — '개인' 또는 '법인'으로 시작해야 함");
            }

            return owner;
        }

        // ─── 3.2.4.3: 구성기업별 선택 ────────────────────────────────────
        // 매년선택 (a,b,c): "2. 매년 선택" 행 기준 오프셋, O(15) = 여/부
        // 5년선택 (d~i):   "3. 5년 선택" 행 기준 오프셋, O(15)=선택연도, Q(17)=취소연도
        // 기타선택 (j):    "6. 기타선택" 행 기준 오프셋, O(15)=선택연도, Q(17)=취소연도
        // 공정가액조정 (k): "k. 공정가액조정" 행 이후 데이터 행, E(5)=사업연도, J(10)=(i)/(ii)
        private static void Map3243(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.3");
            if (rSec < 0) return;

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            var el = ceComp.Elections;

            // ── 매년 선택 (a, b, c) ─────────────────────────────────────────
            var rMaeYear = FindRow(ws, "2. 매년 선택", rSec);
            if (rMaeYear >= 0)
            {
                // a: 간소화 계산 → SimplCalculations
                var vA = ws.Cell(rMaeYear, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(vA))
                {
                    el.SimplCalculations = ParseBool(vA);
                    el.SimplCalculationsSpecified = true;
                }
                // b: 채무면제이익 제외 → Art321
                var vB = ws.Cell(rMaeYear + 1, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(vB))
                {
                    el.Art321 = ParseBool(vB);
                    el.Art321Specified = true;
                }
                // c: 이연법인세비용 미반영(매년) → KArt447C
                var vC = ws.Cell(rMaeYear + 2, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(vC))
                {
                    el.KArt447C = ParseBool(vC);
                    el.KArt447CSpecified = true;
                }
            }

            // ── 5년 선택 (d ~ i) ─────────────────────────────────────────────
            var r5Year = FindRow(ws, "5년 선택", rSec);
            if (r5Year >= 0)
            {
                // d: Art1.5.3 — 제외기업 구성기업 간주 선택
                el.Art153 = ReadYearElectionArt153(ws, r5Year + 1);

                // e: Art3.2.1b — 분산투자지분 배당 포함 선택
                {
                    var (status, ey, ry) = ReadYearElection(ws, r5Year + 2);
                    if (status.HasValue)
                        el.Art321B = new Globe.EtrComputationTypeCeComputationElectionsArt321B
                        {
                            Status = status.Value,
                            ElectionYear = ey ?? default,
                            ElectionYearSpecified = ey.HasValue,
                            RevocationYear = ry ?? default,
                            RevocationYearSpecified = ry.HasValue
                        };
                }

                // f: Art3.2.1c — 환위험 회피수단 외환손익 포함 선택
                {
                    var (status, ey, ry) = ReadYearElection(ws, r5Year + 3);
                    if (status.HasValue)
                        el.Art321C = new Globe.EtrComputationTypeCeComputationElectionsArt321C
                        {
                            Status = status.Value,
                            ElectionYear = ey ?? default,
                            ElectionYearSpecified = ey.HasValue,
                            RevocationYear = ry ?? default,
                            RevocationYearSpecified = ry.HasValue
                        };
                }

                // g: Art7.5[] — 투자기업 투시과세기업 취급 선택
                {
                    var (status, ey, ry) = ReadYearElection(ws, r5Year + 4);
                    if (status.HasValue)
                    {
                        var a75 = new Globe.EtrComputationTypeCeComputationElectionsArt75
                        {
                            Status = status.Value,
                            ElectionYear = ey ?? default,
                            ElectionYearSpecified = ey.HasValue,
                            RevocationYear = ry ?? default,
                            RevocationYearSpecified = ry.HasValue
                        };
                        el.Art75.Add(a75);
                    }
                }

                // h: Art7.6 과세분배방법 적용 선택 — 선택연도/취소연도만 여기서 읽음
                // 상세 데이터(분배금, 공제금액, 지분율, 투자기업TIN)는 3.2.4.5에서 Art76 생성

                // i: Art4.4.7 — 이연법인세비용 미반영(5년) → Art447
                {
                    var (status, ey, ry) = ReadYearElection(ws, r5Year + 6);
                    if (status.HasValue)
                        el.Art447 = new Globe.EtrComputationTypeCeComputationElectionsArt447
                        {
                            Status = status.Value,
                            ElectionYear = ey ?? default,
                            ElectionYearSpecified = ey.HasValue,
                            RevocationYear = ry ?? default,
                            RevocationYearSpecified = ry.HasValue
                        };
                }
            }

            // ── 기타 선택 (j) ─────────────────────────────────────────────────
            var rKita = FindRow(ws, "6. 기타선택", rSec);
            if (rKita >= 0)
            {
                // j: Art4.5.6 — 결손취급특례 적용 선택 (같은 행)
                var (status, ey, ry) = ReadYearElection(ws, rKita);
                if (status.HasValue)
                    el.Art456 = new Globe.EtrComputationTypeCeComputationElectionsArt456
                    {
                        Status = status.Value,
                        ElectionYear = ey ?? default,
                        ElectionYearSpecified = ey.HasValue,
                        RevocationYear = ry ?? default,
                        RevocationYearSpecified = ry.HasValue
                    };
            }

            // ── k: 공정가액조정 선택 → Art634[] ───────────────────────────────
            // B열: "k. 공정가액조정", 이후 +2행부터 데이터 (헤더행 1개 건너뜀)
            // E(5)=FYTriggerEvent(날짜), J(10)=(i)/(ii) → Inclusion.Art634CI/Art634CIi
            var rK = FindRow(ws, "k. 공정가액조정", rSec);
            if (rK >= 0)
            {
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? rK + 10;
                for (int r = rK + 2; r <= System.Math.Min(rK + 10, lastRow); r++)
                {
                    var eyRaw = ws.Cell(r, 5).GetString()?.Trim();  // E: 사업연도
                    var inclRaw = ws.Cell(r, 10).GetString()?.Trim(); // J: (i)/(ii)
                    if (string.IsNullOrEmpty(eyRaw) && string.IsNullOrEmpty(inclRaw)) break;
                    if (!TryParseDate(eyRaw, out var fy)) continue;

                    var incl = new Globe.EtrComputationTypeCeComputationElectionsArt634Inclusion();
                    if (inclRaw == "(i)" || inclRaw == "i") { incl.Art634CI = true; incl.Art634CISpecified = true; }
                    else if (inclRaw == "(ii)" || inclRaw == "ii") { incl.Art634CIi = true; incl.Art634CIiSpecified = true; }

                    el.Art634.Add(new Globe.EtrComputationTypeCeComputationElectionsArt634
                    {
                        FyTriggerEvent = fy,
                        Inclusion = incl
                    });
                }
            }
        }

        // ─── 3.2.4.4: 국제해운소득·결손 제외 → CEComputation.NetGlobeIncome.IntShippingIncome ──
        // rSec+3: 항목2 국제해운 Category(O=15, 쉼표구분), rSec+4: Revenue, rSec+5: Costs, rSec+6: Total
        // rSec+7: 항목6 부수소득 Category(O=15 단일), rSec+8: Revenue, rSec+9: Costs, rSec+10: Total
        // rSec+11: 인건비(PayrollCosts), rSec+12: 유형자산(TangibleAssets), rSec+13: CoveredTaxes
        private static void Map3244(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.4");
            if (rSec < 0) return;

            string V(int offset) => ws.Cell(rSec + offset, 15).GetString()?.Trim();

            // InternationalShipIncome
            var intCatRaw = V(3);
            var intRevenue = V(4);
            var intCosts   = V(5);
            var intTotal   = V(6);

            // QualifiedAncShipIncome
            var ancCatRaw  = V(7);
            var ancRevenue = V(8);
            var ancCosts   = V(9);
            var ancTotal   = V(10);

            // SubstanceExclusion
            var payroll    = V(11);
            var tangible   = V(12);

            // CoveredTaxes
            var covTaxes   = V(13);

            // 모두 빈 값이면 스킵
            if (new[] { intCatRaw, intRevenue, intCosts, intTotal,
                        ancCatRaw, ancRevenue, ancCosts, ancTotal,
                        payroll, tangible, covTaxes }.All(string.IsNullOrEmpty))
                return;

            ceComp.NetGlobeIncome ??= new Globe.EtrComputationTypeCeComputationNetGlobeIncome();
            var ngi = ceComp.NetGlobeIncome;

            ngi.IntShippingIncome = new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncome
            {
                InternationalShipIncome = new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncomeInternationalShipIncome
                {
                    Total   = intTotal,
                    Revenue = intRevenue,
                    Costs   = intCosts,
                },
                QualifiedAncShipIncome = new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncomeQualifiedAncShipIncome
                {
                    Total   = ancTotal,
                    Revenue = ancRevenue,
                    Costs   = ancCosts,
                },
                SubstanceExclusion = new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncomeSubstanceExclusion
                {
                    PayrollCosts  = payroll,
                    TangibleAssets = tangible,
                },
                CoveredTaxes = covTaxes,
            };

            // 국제해운소득 유형(Category) — GIR 코드 추출 (구분자 무관: 쉼표/마침표/공백 등)
            if (!string.IsNullOrEmpty(intCatRaw))
            {
                var intCodes = System.Text.RegularExpressions.Regex.Matches(intCatRaw, @"GIR\d+");
                if (intCodes.Count == 0)
                    errors.Add($"[{fileName}] 3.2.4.4 국제해운소득 유형 파싱 실패: '{intCatRaw}' (GIR2101~GIR2106)");
                foreach (System.Text.RegularExpressions.Match m in intCodes)
                {
                    if (TryParseEnum<Globe.IntShipCategoryEnumType>(m.Value, out var cat))
                        ngi.IntShippingIncome.InternationalShipIncome.Category.Add(cat);
                    else
                        errors.Add($"[{fileName}] 3.2.4.4 국제해운소득 유형 파싱 실패: '{m.Value}' (GIR2101~GIR2106)");
                }
            }

            // 적격국제해운부수소득 유형(Category) — 단일 선택
            if (!string.IsNullOrEmpty(ancCatRaw))
            {
                if (TryParseEnum<Globe.AncShipCategoryEnumType>(ancCatRaw, out var ancCat))
                {
                    ngi.IntShippingIncome.QualifiedAncShipIncome.Category = ancCat;
                }
                else
                    errors.Add($"[{fileName}] 3.2.4.4 적격국제해운부수소득 유형 파싱 실패: '{ancCatRaw}' (GIR2201~GIR2205)");
            }
        }

        // ─── 3.2.4.5: 과세분배방법 적용 선택 관련 정보 → CEComputation.Elections.Art76[] ──
        // 헤더: rSec+2(R156), 데이터 5행: rSec+3~rSec+7
        // B(2)=주주구성기업TIN, E(5)=투자기업TIN, G(7)=ActualDeemedDist,
        // K(11)=LocalCreditableTaxGross, N(14)=ShareOfUndistNetGlobeInc
        // Status/ElectionYear/RevocationYear: 3.2.4.3 h행(r5Year+5)에서 읽음
        private static void Map3245(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.5");
            if (rSec < 0) return;

            // 3.2.4.3 h행에서 Status/ElectionYear/RevocationYear 읽기
            bool art76Status = true;
            System.DateTime art76Ey = default;
            System.DateTime? art76Ry = null;
            var r3243 = FindRow(ws, "3.2.4.3");
            if (r3243 >= 0)
            {
                var r5Year = FindRow(ws, "5년 선택", r3243);
                if (r5Year >= 0)
                {
                    var (st, ey, ry) = ReadYearElection(ws, r5Year + 5);
                    if (st.HasValue) art76Status = st.Value;
                    if (ey.HasValue) art76Ey = ey.Value;
                    art76Ry = ry;
                }
            }

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            var el = ceComp.Elections;

            // 데이터 5행: rSec+3 ~ rSec+7
            for (int offset = 3; offset <= 7; offset++)
            {
                var r = rSec + offset;
                var eRaw = ws.Cell(r, 5).GetString()?.Trim();  // E: 투자기업 TIN
                var gRaw = ws.Cell(r, 7).GetString()?.Trim();  // G: ActualDeemedDist
                var kRaw = ws.Cell(r, 11).GetString()?.Trim(); // K: LocalCreditableTaxGross
                var nRaw = ws.Cell(r, 14).GetString()?.Trim(); // N: ShareOfUndistNetGlobeInc

                if (string.IsNullOrEmpty(eRaw) && string.IsNullOrEmpty(gRaw)) continue;

                var a76 = new Globe.EtrComputationTypeCeComputationElectionsArt76
                {
                    Status = art76Status,
                    ElectionYear = art76Ey,
                    RevocationYear = art76Ry ?? default,
                    RevocationYearSpecified = art76Ry.HasValue,
                    ActualDeemedDist = gRaw,
                    LocalCreditableTaxGross = kRaw,
                };

                if (!string.IsNullOrEmpty(nRaw))
                {
                    if (decimal.TryParse(nRaw, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var share))
                        a76.ShareOfUndistNetGlobeInc = share;
                    else
                        errors.Add($"[{fileName}] 3.2.4.5 지분율 파싱 실패: '{nRaw}' (0~1 소수)");
                }

                if (!string.IsNullOrEmpty(eRaw))
                {
                    var invTin = ParseTin(eRaw);
                    if (invTin != null)
                        a76.InvestmentEntityTin = invTin;
                    else
                        errors.Add($"[{fileName}] 3.2.4.5 투자기업 TIN 파싱 실패: '{eRaw}'");
                }

                el.Art76.Add(a76);
            }
        }

        // ─── 3.2.4.6: 그 밖의 회계기준 → CEComputation.OtherFas ──────────────
        // 헤더: rSec+1(R163), 데이터 행: rSec+2 이후 (B가 빈 행까지)
        // B(2)=CE TIN, K(11)=OtherFas (적용 인정회계기준/공인회계기준)
        // ceComp.Tin.Value와 일치하는 행의 K값을 사용, 없으면 첫 번째 비어있지 않은 K값
        private static void Map3246(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.6");
            if (rSec < 0) return;

            var currentTin = ceComp.Tin?.Value?.Trim();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            string matched = null;
            string firstFound = null;

            for (int r = rSec + 2; r <= lastRow; r++)
            {
                var bRaw = ws.Cell(r, 2).GetString()?.Trim();
                var kRaw = ws.Cell(r, 11).GetString()?.Trim();

                if (string.IsNullOrEmpty(bRaw) && string.IsNullOrEmpty(kRaw)) break;
                if (string.IsNullOrEmpty(kRaw)) continue;

                if (firstFound == null) firstFound = kRaw;

                if (!string.IsNullOrEmpty(currentTin) && !string.IsNullOrEmpty(bRaw)
                    && bRaw.Contains(currentTin))
                {
                    matched = kRaw;
                    break;
                }
            }

            var result = matched ?? firstFound;
            if (!string.IsNullOrEmpty(result))
                ceComp.OtherFas = result;
        }

        // ─── 3.4.1: 소득산입규칙(IIR) 적용 → JurisdictionSection.LowTaxJurisdiction ──
        // 1. a,b,c → 주 시트 O(15)열 / 2+3 모기업 테이블 → "추가세액 첨부" 시트 B~H
        private static void Map341(
            IXLWorksheet ws,
            Globe.JurisdictionSectionType js,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.4.1");
            if (rSec < 0) return;

            // 1. 추가세액 배분 그룹기업 (LTCE) a,b,c → O(15)열
            // B열 "1. 추가세액 배분 그룹기업" 헤더 위치 = a 행 (merged cell 상단)
            var rLtce = FindRow(ws, "1. 추가세액 배분", rSec);
            if (rLtce < 0) return;

            var tinRaw = ws.Cell(rLtce,     15).GetString()?.Trim(); // a: TIN
            var ngiRaw = ws.Cell(rLtce + 1, 15).GetString()?.Trim(); // b: NetGlobeIncome
            var tutRaw = ws.Cell(rLtce + 2, 15).GetString()?.Trim(); // c: TopUpTax
            if (string.IsNullOrEmpty(tinRaw)) return;

            // 첨부 시트 탐색: 3.4.1 구간의 O열에서 시트 참조 읽기
            IXLWorksheet attachWs = null;
            var rEnd = FindRow(ws, "3.4.2", rSec + 1);
            if (rEnd < 0) rEnd = ws.LastRowUsed()?.RowNumber() ?? 500;
            for (int r = rLtce; r <= rEnd; r++)
            {
                var oVal = ws.Cell(r, 15).GetString()?.Trim();
                if (string.IsNullOrEmpty(oVal) || oVal == tinRaw || oVal == ngiRaw || oVal == tutRaw) continue;
                if (ws.Workbook.TryGetWorksheet(oVal, out attachWs)) break;
            }
            if (attachWs == null) // 폴백: "첨부" 포함 시트 검색
                attachWs = ws.Workbook.Worksheets
                    .FirstOrDefault(s => s.Name != ws.Name && s.Name.Contains("첨부"));

            // IIR 엔트리 생성
            var iir = new Globe.LowTaxJurisdictionTypeLtceIir
            {
                NetGlobeIncome = string.IsNullOrEmpty(ngiRaw) ? null : ngiRaw,
                TopUpTax = string.IsNullOrEmpty(tutRaw) ? "0" : tutRaw
            };

            // 모기업 테이블 (첨부 시트) → ParentEntity[]
            if (attachWs != null)
            {
                var lastRow = attachWs.LastRowUsed()?.RowNumber() ?? 100;
                // 헤더 행 건너뛰기: B열에 "모기업" 또는 "납세자번호" 포함 행 이후부터
                int dataStart = 1;
                for (int r = 1; r <= System.Math.Min(lastRow, 10); r++)
                {
                    var hdr = attachWs.Cell(r, 2).GetString();
                    if (hdr != null && (hdr.Contains("모기업") || hdr.Contains("납세자번호")))
                    { dataStart = r + 1; break; }
                }
                for (int r = dataStart; r <= lastRow; r++)
                {
                    var bRaw = attachWs.Cell(r, 2).GetString()?.Trim(); // B: 모기업 TIN
                    var cRaw = attachWs.Cell(r, 3).GetString()?.Trim(); // C: 소재지국
                    var dRaw = attachWs.Cell(r, 4).GetString()?.Trim(); // D: OtherOwnershipAllocation
                    var eRaw = attachWs.Cell(r, 5).GetString()?.Trim(); // E: InclusionRatio
                    var fRaw = attachWs.Cell(r, 6).GetString()?.Trim(); // F: TopUpTaxShare
                    var gRaw = attachWs.Cell(r, 7).GetString()?.Trim(); // G: IirOffSet
                    var hRaw = attachWs.Cell(r, 8).GetString()?.Trim(); // H: TopUpTax

                    if (string.IsNullOrEmpty(bRaw) && string.IsNullOrEmpty(fRaw)) break;
                    if (string.IsNullOrEmpty(bRaw)) continue;

                    Globe.CountryCodeType cc = default;
                    if (!string.IsNullOrEmpty(cRaw))
                        System.Enum.TryParse(cRaw, true, out cc);

                    decimal.TryParse(eRaw,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out var inclusionRatio);

                    iir.ParentEntity.Add(new Globe.LowTaxJurisdictionTypeLtceIirParentEntity
                    {
                        Tin                    = ParseTin(bRaw),
                        ResCountryCode         = cc,
                        OtherOwnershipAllocation = dRaw ?? "0",
                        InclusionRatio         = inclusionRatio,
                        TopUpTaxShare          = fRaw ?? "0",
                        IirOffSet              = gRaw ?? "0",
                        TopUpTax               = hRaw ?? "0"
                    });
                }
            }

            var ltce = new Globe.LowTaxJurisdictionTypeLtce { Tin = ParseTin(tinRaw) };
            ltce.Iir.Add(iir);

            js.LowTaxJurisdiction ??= new Globe.LowTaxJurisdictionType { TopUpTaxAmount = "0" };
            js.LowTaxJurisdiction.Ltce.Add(ltce);
        }

        // ─── 3.4.2: 소득산입보완규칙(UTPR) → LowTaxJurisdiction.Utpr ─────────
        // 1: 납세자번호(XML 미포함), 2: Article2.5.1TopUpTax, 3: TotalUTPRTopUpTax — 모두 O(15)열
        private static void Map342(
            IXLWorksheet ws,
            Globe.JurisdictionSectionType js,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.4.2");
            if (rSec < 0) return;

            // 3.4.2.2: "제73조제3항제2호에 따라" → Article251TopUpTax
            var r2 = FindRow(ws, "3항제2호", rSec);
            if (r2 < 0) r2 = FindRow(ws, "납부액을 차감", rSec); // 폴백
            // 3.4.2.3: 소득산입보완규칙 추가세액 합계 → TotalUtprTopUpTax
            var r3 = FindRow(ws, "소득산입보완규칙 추가세액 합계", rSec);

            var art251   = r2 > 0 ? ws.Cell(r2, 15).GetString()?.Trim() : null;
            var totalUtpr = r3 > 0 ? ws.Cell(r3, 15).GetString()?.Trim() : null;
            if (string.IsNullOrEmpty(art251) && string.IsNullOrEmpty(totalUtpr)) return;

            js.LowTaxJurisdiction ??= new Globe.LowTaxJurisdictionType { TopUpTaxAmount = "0" };
            js.LowTaxJurisdiction.Utpr ??= new Globe.LowTaxJurisdictionTypeUtpr();
            js.LowTaxJurisdiction.Utpr.UtprCalculation = new Globe.LowTaxJurisdictionTypeUtprUtprCalculation
            {
                Article251TopUpTax = art251   ?? "0",
                TotalUtprTopUpTax  = totalUtpr ?? "0"
            };
        }

        // 선택 사업연도(O=15)/취소 사업연도(Q=17) 읽기 헬퍼
        // Status: 선택연도가 있으면 true, 없고 취소연도만 있으면 false
        private static (bool? status, System.DateTime? ey, System.DateTime? ry) ReadYearElection(IXLWorksheet ws, int row)
        {
            var eyRaw = ws.Cell(row, 15).GetString()?.Trim(); // O
            var ryRaw = ws.Cell(row, 17).GetString()?.Trim(); // Q
            if (string.IsNullOrEmpty(eyRaw) && string.IsNullOrEmpty(ryRaw)) return (null, null, null);

            TryParseDate(eyRaw, out var ey);
            TryParseDate(ryRaw, out var ry);
            bool status = !string.IsNullOrEmpty(eyRaw);
            return (status, ey == default ? (System.DateTime?)null : ey, ry == default ? (System.DateTime?)null : ry);
        }

        // Art153 전용
        private static Globe.EtrComputationTypeCeComputationElectionsArt153 ReadYearElectionArt153(IXLWorksheet ws, int row)
        {
            var (status, ey, ry) = ReadYearElection(ws, row);
            if (!status.HasValue) return null;
            return new Globe.EtrComputationTypeCeComputationElectionsArt153
            {
                Status = status.Value,
                ElectionYear = ey ?? default,
                ElectionYearSpecified = ey.HasValue,
                RevocationYear = ry ?? default,
                RevocationYearSpecified = ry.HasValue
            };
        }

        // B열에서 contains 텍스트를 포함하는 첫 번째 행 반환 (-1 = 없음)
        private static int FindRow(IXLWorksheet ws, string contains, int fromRow = 1)
        {
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            for (int r = fromRow; r <= lastRow; r++)
            {
                var v = ws.Cell(r, 2).GetString();
                if (v != null && v.Contains(contains))
                    return r;
            }
            return -1;
        }
    }
}
