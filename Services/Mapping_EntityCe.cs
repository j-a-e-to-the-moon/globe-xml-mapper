using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// main_template_newest.xlsx "구성기업 계산" 시트 → CEComputation 매핑.
    /// </summary>
    public class Mapping_EntityCe : MappingBase
    {
        public Mapping_EntityCe()
            : base(null) { }

        // 현재 처리 중인 블록 범위 (FindRow 스코프)
        private int _blockStart = 1;
        private int _blockEnd = -1;
        // 시트 전역에 1번만 존재하는 헤더 캐시 (다중 CE 블록 시 모든 CE에 동일 적용)
        private int _globalSimplCalcRow = -1;

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            // 구성기업 시트 로드
            var groupMap = EntityGroupMap.Load(ws.Workbook, errors);

            // 시트 최상단(블록 외부)에 1번만 있는 헤더 위치 캐시
            // FindRow는 _blockStart 스코프라 두 번째 이상 블록에서는 못 찾기 때문
            _globalSimplCalcRow = FindRowGlobal(ws, "1. 귀 다국적기업그룹은");

            // 세로 스택된 entity 블록 모두 찾기
            var blockStarts = FindAllEntityBlockStarts(ws);
            if (blockStarts.Count == 0)
                return;

            var lastUsedRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            for (int i = 0; i < blockStarts.Count; i++)
            {
                _blockStart = blockStarts[i];
                _blockEnd = (i + 1 < blockStarts.Count) ? blockStarts[i + 1] - 1 : lastUsedRow;
                MapOneEntity(ws, globe, groupMap, errors, fileName);
            }

            _blockStart = 1;
            _blockEnd = -1;
            _globalSimplCalcRow = -1;
        }

        // 블록 스코프 무시하고 시트 전역에서 헤더 검색.
        private static int FindRowGlobal(IXLWorksheet ws, string contains)
        {
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            for (int r = 1; r <= lastRow; r++)
            {
                var v = ws.Cell(r, 2).GetString() ?? "";
                if (v.Contains(contains)) return r;
            }
            return -1;
        }

        /// <summary>
        /// 전체 시트에서 entity 블록 시작 행 모두 찾기.
        /// 마커: "3.2.4 구성기업" — 각 entity 블록 최상위 헤더.
        /// 기존 "1. 구성기업 또는 공동기업그룹 기업의 납세자번호"는 한 블록 내에서
        /// 3.2.4.1(a) / 3.2.4.2(a) / 3.2.4.2(c) 섹션마다 중복 등장해 블록이 과분할되는 문제가 있었음.
        /// </summary>
        private static List<int> FindAllEntityBlockStarts(IXLWorksheet ws)
        {
            var result = new List<int>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            for (int r = 1; r <= lastRow; r++)
            {
                var v = ws.Cell(r, 2).GetString() ?? "";
                if (v.Contains("3.2.4 구성기업"))
                    result.Add(r);
            }
            return result;
        }

        private void MapOneEntity(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            EntityGroupMap groupMap,
            List<string> errors,
            string fileName
        )
        {
            // ── CE TIN 식별 (3.2.4.1(a) 납세자번호 M열) ─────────
            var rTin = FindRow(ws, "1. 구성기업 또는 공동기업그룹 기업의 납세자번호");
            if (rTin < 0)
            {
                errors.Add($"[{fileName}] [블록 {_blockStart}~{_blockEnd}] CE TIN 항목 없음");
                return;
            }

            var ceTinRaw = ws.Cell(rTin, 13).GetString()?.Trim(); // M열 (M15)
            if (string.IsNullOrEmpty(ceTinRaw))
            {
                // 빈 블록 silent skip — 템플릿 헤더만 있고 사용자가 아직 CE 입력 안 한 경우.
                // 실제 CE를 보고할 때는 이 필드가 채워져야 함 (검증 단계에서 catch).
                return;
            }

            var ceTin = ParseTin(ceTinRaw);

            // ── 기업매핑에서 (국가, 하위그룹 TIN) 조회 ─────────
            Globe.CountryCodeType lookupCode;
            string subGroupTin = null;
            if (groupMap.TryGet(ceTin.Value, out var mapEntry))
            {
                if (mapEntry.Country.HasValue)
                {
                    lookupCode = mapEntry.Country.Value;
                    subGroupTin = mapEntry.SubGroupTin;
                }
                else
                {
                    errors.Add(
                        $"[{fileName}] [블록 R{_blockStart}] 구성기업 시트에 TIN '{ceTin.Value}' 있으나 국가 미지정"
                    );
                    return;
                }
            }
            else if (ceTin.IssuedBySpecified)
            {
                // 기업매핑에 없으면 TIN의 발급국가로 폴백
                lookupCode = ceTin.IssuedBy;
            }
            else
            {
                errors.Add(
                    $"[{fileName}] [블록 R{_blockStart}] CE TIN '{ceTin.Value}' 이 구성기업 시트에 없음 + 발급국가도 없음"
                );
                return;
            }

            var js = globe.GlobeBody.JurisdictionSection.FirstOrDefault(s =>
                s.Jurisdiction == lookupCode
            );
            if (js == null)
            {
                errors.Add(
                    $"[{fileName}] [블록 R{_blockStart}] JurisdictionSection({lookupCode}) 없음 — '국가별 계산' 시트에 해당 합산단위 블록 필요"
                );
                return;
            }

            // ── ETR 찾기: SubGroup TIN 일치 우선, 없으면 첫 번째 ─────────
            Globe.EtrType etr;
            if (!string.IsNullOrEmpty(subGroupTin))
                etr =
                    js.GLoBeTax.Etr.FirstOrDefault(e => e.SubGroup?.Tin?.Value == subGroupTin)
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
            var ceComp = etr.EtrStatus.EtrComputation.CeComputation.FirstOrDefault(c =>
                c.Tin?.Value == ceTin.Value
            );
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

            // ── 3.4.1/3.4.2 추가세액 — 구성기업 계산 시트 내부에 통합됨 ──────
            Map341(ws, js, errors, fileName);
            Map342(ws, js, errors, fileName);

            // 후처리: 빈 하위 객체 정리 (데이터 없이 생성된 Elections 등)
            CleanupEmptyCeComp(ceComp);
        }

        // ─── 3.2.4(a): row에서 K열 여/부 → Elections.SimplCalculations ──
        // "1. 귀 다국적기업그룹은…" 헤더는 시트 최상단에 1번만 존재.
        // _globalSimplCalcRow 캐시로 다중 CE 블록에서도 동일 값 적용.
        private void Map324a(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            if (_globalSimplCalcRow < 0)
                return;
            var v = ws.Cell(_globalSimplCalcRow, 15).GetString()?.Trim(); // O열 (O4)
            if (string.IsNullOrEmpty(v))
                return;

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            ceComp.Elections.SimplCalculations = ParseBool(v);
            ceComp.Elections.SimplCalculationsSpecified = true;
        }

        // ─── 3.2.4(b): 헤더 다음 행들에서 B=GroupTIN, K=EntityTINs ──────
        // 헤더행: B="1. 연결납세그룹(납세자번호)", K="2. 해당 연결납세그룹에 포함된 기업..."
        // 데이터행: B=GroupTIN ("값,유형,발급국가"), K=EntityTIN (행마다 1개)
        private void Map324b(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "1. 연결납세그룹(납세자번호)");
            if (rHdr < 0)
                return;

            string groupTinRaw = null;
            var entityTins = new System.Collections.Generic.List<Globe.TinType>();

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 10;
            for (int r = rHdr + 1; r <= System.Math.Min(rHdr + 20, lastRow); r++)
            {
                var bVal = ws.Cell(r, 2).GetString()?.Trim(); // B: GroupTIN
                var kVal = ws.Cell(r, 11).GetString()?.Trim(); // K: EntityTIN

                if (string.IsNullOrEmpty(bVal) && string.IsNullOrEmpty(kVal))
                    break;

                if (!string.IsNullOrEmpty(bVal) && groupTinRaw == null)
                    groupTinRaw = bVal;

                if (!string.IsNullOrEmpty(kVal))
                    entityTins.Add(ParseTin(kVal));
            }

            if (groupTinRaw == null)
                return;

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            var ar = new Globe.EtrComputationTypeCeComputationElectionsAggregatedReporting
            {
                TaxConsolGroupTin = ParseTin(groupTinRaw),
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
            Globe.CurrentAdjustedTaxEnumType.Gir2401,
            Globe.CurrentAdjustedTaxEnumType.Gir2402,
            Globe.CurrentAdjustedTaxEnumType.Gir2403,
            Globe.CurrentAdjustedTaxEnumType.Gir2404,
            Globe.CurrentAdjustedTaxEnumType.Gir2405,
            Globe.CurrentAdjustedTaxEnumType.Gir2406,
            Globe.CurrentAdjustedTaxEnumType.Gir2407,
            Globe.CurrentAdjustedTaxEnumType.Gir2408,
            Globe.CurrentAdjustedTaxEnumType.Gir2409,
            Globe.CurrentAdjustedTaxEnumType.Gir2410,
            Globe.CurrentAdjustedTaxEnumType.Gir2411,
            Globe.CurrentAdjustedTaxEnumType.Gir2412,
            Globe.CurrentAdjustedTaxEnumType.Gir2413,
            Globe.CurrentAdjustedTaxEnumType.Gir2414,
            Globe.CurrentAdjustedTaxEnumType.Gir2415,
            Globe.CurrentAdjustedTaxEnumType.Gir2416,
            Globe.CurrentAdjustedTaxEnumType.Gir2417,
        };

        private void Map3242a(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.2 조정대상조세");
            if (rSec < 0)
                return;

            // 2. 배분 후 당기법인세비용 (O열) → AdjustedIncomeTax.Total
            var rIncomeTax = FindRow(ws, "배분 후 당기법인세비용", rSec);
            if (rIncomeTax >= 0)
            {
                var v = ws.Cell(rIncomeTax, 15).GetString()?.Trim(); // O열 (O69)
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.AdjustedIncomeTax ??=
                        new Globe.EtrComputationTypeCeComputationAdjustedIncomeTax();
                    ceComp.AdjustedIncomeTax.Total = v;
                }
            }

            // 3. 조정사항 (a)~(q) 17개 → AdjustedCoveredTax.Adjustments (GIR2401~2417)
            // O(15)=가산액, Q(17)=차감액
            // 실제 값이 하나라도 있을 때만 AdjustedCoveredTax 생성 (빈 태그 방지)
            var rAdjHdr = FindRow(ws, "3. 조정사항", rSec);
            if (rAdjHdr >= 0)
            {
                for (int i = 0; i < TaxAdjustmentItems.Length; i++)
                {
                    var adds = ws.Cell(rAdjHdr + 1 + i, 15).GetString()?.Trim(); // O: 가산액
                    var reds = ws.Cell(rAdjHdr + 1 + i, 17).GetString()?.Trim(); // Q: 차감액
                    if (!HasNonZeroValue(adds) && !HasNonZeroValue(reds))
                        continue;

                    ceComp.AdjustedCoveredTax ??=
                        new Globe.EtrComputationTypeCeComputationAdjustedCoveredTax();

                    var adj = new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxAdjustments
                    {
                        AdjustmentItem = TaxAdjustmentItems[i],
                    };
                    if (HasNonZeroValue(adds))
                        adj.Amount.Add(adds);
                    if (HasNonZeroValue(reds))
                        adj.Amount.Add("-" + reds);
                    ceComp.AdjustedCoveredTax.Adjustments.Add(adj);
                }
            }

            // 4. 조정대상조세 (O열) → AdjustedCoveredTax.Total
            var rTotal = FindRow(ws, "4. 조정대상조세", rSec);
            if (rTotal >= 0)
            {
                var v = ws.Cell(rTotal, 15).GetString()?.Trim(); // O열 (O88)
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.AdjustedCoveredTax ??=
                        new Globe.EtrComputationTypeCeComputationAdjustedCoveredTax();
                    ceComp.AdjustedCoveredTax.Total = v;
                }
            }
        }

        // ─── 3.2.4.2(b): 대상조세 국가간 배분 → AdjustedIncomeTax.IncomeTax + CrossAllocation[] ──
        // 헤더행: "대상조세 국가간 배분", 컬럼헤더 +1행, 데이터 +2행~ (행추가 가능)
        // C(3)=배분전(IncomeTax, 첫 행만), E(5)=Basis, G(7)=OtherTIN, J(10)=ResCountry
        // L(12)=Additions, N(14)=Reductions
        private void Map3242b(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "대상조세 국가간 배분");
            if (rHdr < 0)
                return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            bool firstRow = true;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 30, lastRow); r++)
            {
                var basisRaw = ws.Cell(r, 5).GetString()?.Trim(); // E: 조정 근거
                var otherTinRaw = ws.Cell(r, 7).GetString()?.Trim(); // G: 상대TIN
                var adds = ws.Cell(r, 12).GetString()?.Trim(); // L: 가산액
                var reds = ws.Cell(r, 14).GetString()?.Trim(); // N: 차감액

                if (
                    string.IsNullOrEmpty(basisRaw)
                    && string.IsNullOrEmpty(otherTinRaw)
                    && string.IsNullOrEmpty(adds)
                    && string.IsNullOrEmpty(reds)
                )
                    break;

                // 첫 데이터행 C(3): 배분 전 대상조세 → AdjustedIncomeTax.IncomeTax
                if (firstRow)
                {
                    firstRow = false;
                    var preTax = ws.Cell(r, 3).GetString()?.Trim();
                    if (!string.IsNullOrEmpty(preTax))
                    {
                        ceComp.AdjustedIncomeTax ??=
                            new Globe.EtrComputationTypeCeComputationAdjustedIncomeTax();
                        ceComp.AdjustedIncomeTax.IncomeTax = preTax;
                    }
                }

                if (string.IsNullOrEmpty(basisRaw))
                    continue;

                if (!TryParseEnum<Globe.AdjustedBasisEnumType>(basisRaw, out var basis))
                {
                    errors.Add(
                        $"[{fileName}] 3.2.4.2(b) 조정 근거 '{basisRaw}' 파싱 실패 — GIR2201~GIR2205 중 하나여야 함"
                    );
                    continue;
                }

                var item =
                    new Globe.EtrComputationTypeCeComputationAdjustedIncomeTaxCrossAllocation();
                item.Basis.Add(basis);
                if (!string.IsNullOrEmpty(otherTinRaw))
                    item.OtherTin = ParseTin(otherTinRaw);
                if (HasNonZeroValue(adds))
                    item.Additions = adds;
                if (HasNonZeroValue(reds))
                    item.Reductions = reds;

                var ccRaw = ws.Cell(r, 10).GetString()?.Trim(); // J: 소재지국 (J4)
                if (
                    !string.IsNullOrEmpty(ccRaw)
                    && TryParseEnum<Globe.CountryCodeType>(ccRaw, out var cc)
                )
                {
                    item.ResCountryCode = cc;
                }

                ceComp.AdjustedIncomeTax ??=
                    new Globe.EtrComputationTypeCeComputationAdjustedIncomeTax();
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

        private void Map3242c(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "(c) 이연법인세");
            if (rSec < 0)
                return;

            // 2. 이연법인세비용 금액(글로벌최저한세 기준) O열 → DeferTaxExpense
            var rExpense = FindRow(ws, "이연법인세비용 금액", rSec);
            string deferExpenseVal = null;
            if (rExpense >= 0)
                deferExpenseVal = ws.Cell(rExpense, 15).GetString()?.Trim(); // O96

            // 3. 조정항목 (a)~(p) 16개: O(15)=가산액, Q(17)=차감액
            var rAdjHdr = FindRow(ws, "이연법인세비용의 조정", rSec);
            var adjItems = new System.Collections.Generic.Dictionary<
                Globe.DeferredAdjustedTaxEnumType,
                Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment
            >();

            if (rAdjHdr >= 0)
            {
                for (int i = 0; i < DeferredTaxAdjItems.Length; i++)
                {
                    var adds = ws.Cell(rAdjHdr + 1 + i, 15).GetString()?.Trim(); // O: 가산액
                    var reds = ws.Cell(rAdjHdr + 1 + i, 17).GetString()?.Trim(); // Q: 차감액
                    if (!HasNonZeroValue(adds) && !HasNonZeroValue(reds))
                        continue;

                    var adj =
                        new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment
                        {
                            AdjustmentItem = DeferredTaxAdjItems[i],
                        };
                    if (HasNonZeroValue(adds))
                        adj.Amount.Add(adds);
                    if (HasNonZeroValue(reds))
                        adj.Amount.Add("-" + reds);
                    adjItems[DeferredTaxAdjItems[i]] = adj;
                }
            }

            // 4. 낮은 세율 재계산 차이금액 → GIR2511 항목의 Recast.Lower
            var rRecastDown = FindRow(ws, "최저한세율보다 낮은 세율", rSec);
            if (rRecastDown >= 0)
            {
                var v = ws.Cell(rRecastDown, 15).GetString()?.Trim(); // O114
                if (!string.IsNullOrEmpty(v))
                {
                    if (
                        !adjItems.TryGetValue(
                            Globe.DeferredAdjustedTaxEnumType.Gir2511,
                            out var adj2511
                        )
                    )
                    {
                        adj2511 =
                            new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment
                            {
                                AdjustmentItem = Globe.DeferredAdjustedTaxEnumType.Gir2511,
                            };
                        adjItems[Globe.DeferredAdjustedTaxEnumType.Gir2511] = adj2511;
                    }
                    adj2511.Recast ??=
                        new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentRecast();
                    adj2511.Recast.Lower = v;
                }
            }

            // 5. 높은 세율 재계산 차이금액 → GIR2512 항목의 Recast.Higher
            var rRecastUp = FindRow(ws, "최저한세율보다 높은 세율", rSec);
            if (rRecastUp >= 0)
            {
                var v = ws.Cell(rRecastUp, 15).GetString()?.Trim(); // O115
                if (!string.IsNullOrEmpty(v))
                {
                    if (
                        !adjItems.TryGetValue(
                            Globe.DeferredAdjustedTaxEnumType.Gir2512,
                            out var adj2512
                        )
                    )
                    {
                        adj2512 =
                            new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustment
                            {
                                AdjustmentItem = Globe.DeferredAdjustedTaxEnumType.Gir2512,
                            };
                        adjItems[Globe.DeferredAdjustedTaxEnumType.Gir2512] = adj2512;
                    }
                    adj2512.Recast ??=
                        new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmtAdjustmentRecast();
                    adj2512.Recast.Higher = v;
                }
            }

            // 6. 총이연법인세조정금액 → Total
            var rTotal = FindRow(ws, "총이연법인세조정금액", rSec);
            string totalVal = null;
            if (rTotal >= 0)
                totalVal = ws.Cell(rTotal, 15).GetString()?.Trim(); // O86

            // 의미 있는 값이 있을 때만 DeferTaxAdjustAmt 생성
            // Total=0만 있고 expense/adjustments 없으면 빈 구조체이므로 스킵
            if (
                HasNonZeroValue(deferExpenseVal)
                || adjItems.Count > 0
                || HasNonZeroValue(totalVal)
            )
            {
                ceComp.AdjustedCoveredTax ??=
                    new Globe.EtrComputationTypeCeComputationAdjustedCoveredTax();
                var dta =
                    ceComp.AdjustedCoveredTax.DeferTaxAdjustAmt
                    ?? new Globe.EtrComputationTypeCeComputationAdjustedCoveredTaxDeferTaxAdjustAmt();
                ceComp.AdjustedCoveredTax.DeferTaxAdjustAmt = dta;

                if (!string.IsNullOrEmpty(deferExpenseVal))
                    dta.DeferTaxExpense = deferExpenseVal;
                if (!string.IsNullOrEmpty(totalVal))
                    dta.Total = totalVal;

                // Adjustment[]는 enum 순서(GIR2501~2516)대로 추가
                foreach (var item in DeferredTaxAdjItems)
                    if (adjItems.TryGetValue(item, out var a))
                        dta.Adjustment.Add(a);
            }
        }

        // ─── 3.2.4.1(a): FANIL → AdjustedFanil.Fanil
        //                  (a)~(z) 26개 조정항목 → NetGlobeIncome.Adjustments[]
        //                  글로벌최저한세소득결손 합계 → NetGlobeIncome.Total
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

        private void Map3241a(
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
                var v = ws.Cell(rFanil, 13).GetString()?.Trim(); // M16
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.AdjustedFanil ??=
                        new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                    ceComp.AdjustedFanil.Total = v;
                }
            }

            // (a)~(z) 26개 조정항목 (rows 18~43) → NetGlobeIncome.Adjustments[]
            // 헤더행 "3. 조정사항" 바로 다음 행부터 26행이 고정 구조
            // 실제 값이 하나라도 있을 때만 NetGlobeIncome 객체 생성 (빈 태그 방지)
            var rAdjHdr = FindRow(ws, "3. 조정사항");
            if (rAdjHdr >= 0)
            {
                for (int i = 0; i < AdjustmentItems.Length; i++)
                {
                    var adds = ws.Cell(rAdjHdr + 1 + i, 13).GetString()?.Trim(); // M: 가산액
                    var reds = ws.Cell(rAdjHdr + 1 + i, 16).GetString()?.Trim(); // P: 차감액
                    if (!HasNonZeroValue(adds) && !HasNonZeroValue(reds))
                        continue;

                    ceComp.NetGlobeIncome ??=
                        new Globe.EtrComputationTypeCeComputationNetGlobeIncome();

                    var adj = new Globe.EtrComputationTypeCeComputationNetGlobeIncomeAdjustments
                    {
                        AdjustmentItem = AdjustmentItems[i],
                    };
                    if (HasNonZeroValue(adds))
                        adj.Amount.Add(adds);
                    if (HasNonZeroValue(reds))
                        adj.Amount.Add("-" + reds);
                    ceComp.NetGlobeIncome.Adjustments.Add(adj);
                }
            }

            // 글로벌최저한세소득결손 합계 (row 44) → NetGlobeIncome.Total
            var rTotal = FindRow(ws, "4. 구성기업 또는 공동기업그룹 기업의 글로벌최저한세소득");
            if (rTotal >= 0)
            {
                var v = ws.Cell(rTotal, 13).GetString()?.Trim(); // M44
                if (!string.IsNullOrEmpty(v))
                {
                    ceComp.NetGlobeIncome ??=
                        new Globe.EtrComputationTypeCeComputationNetGlobeIncome();
                    ceComp.NetGlobeIncome.Total = v;
                }
            }
        }

        // ─── 3.2.4.1(b): 본점·고정사업장 간 배분 → MainEntityPEandFte ──
        // 헤더행: "(b) 본점과 고정사업장", 컬럼헤더 +1행, 데이터 +2행~ (행추가 가능)
        // E(5)=Basis(GIR1701~1704), G(7)=OtherTIN, J(10)=ResCountry, L(12)=Additions, N(14)=Reductions
        // 첫 데이터행 C(3)="2. 조정 전 회계상 순손익" 값 → AdjustedFanil.Fanil
        // (b) 섹션 없거나 C(3) 비어있으면 M16 "배분 후 회계상 순손익" 으로 폴백
        private void Map3241b(
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
                fanilVal = ws.Cell(rHdr + 2, 3).GetString()?.Trim(); // C열: 조정 전 회계상 순손익 (C48)
            if (string.IsNullOrEmpty(fanilVal))
            {
                var rM16 = FindRow(ws, "배분 후 회계상 순손익");
                if (rM16 >= 0)
                    fanilVal = ws.Cell(rM16, 13).GetString()?.Trim(); // M16
            }
            if (!string.IsNullOrEmpty(fanilVal))
            {
                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Fanil = fanilVal;
            }

            if (rHdr < 0)
                return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 15, lastRow); r++)
            {
                var basisRaw = ws.Cell(r, 5).GetString()?.Trim(); // E: 조정근거
                var otherTinRaw = ws.Cell(r, 7).GetString()?.Trim(); // G: 상대TIN
                var adds = ws.Cell(r, 12).GetString()?.Trim(); // L: 가산액
                var reds = ws.Cell(r, 14).GetString()?.Trim(); // N: 차감액

                if (
                    string.IsNullOrEmpty(basisRaw)
                    && string.IsNullOrEmpty(otherTinRaw)
                    && string.IsNullOrEmpty(adds)
                    && string.IsNullOrEmpty(reds)
                )
                    break;
                if (string.IsNullOrEmpty(basisRaw))
                    continue;

                if (!TryParseEnum<Globe.MainEntityPEandFteBasisEnumType>(basisRaw, out var basis))
                {
                    errors.Add(
                        $"[{fileName}] 3.2.4.1(b) 조정근거 '{basisRaw}' 파싱 실패 — GIR1701~GIR1704 중 하나여야 함"
                    );
                    continue;
                }

                var item =
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentMainEntityPEandFte
                    {
                        Basis = basis,
                        OtherTin = ParseTin(otherTinRaw ?? ""),
                        Additions = adds ?? "0",
                        Reductions = reds ?? "0",
                    };

                var ccRaw = ws.Cell(r, 10).GetString()?.Trim(); // J: 소재지국 (J4)
                if (
                    !string.IsNullOrEmpty(ccRaw)
                    && TryParseEnum<Globe.CountryCodeType>(ccRaw, out var cc)
                )
                {
                    item.ResCountryCode = cc;
                    item.ResCountryCodeSpecified = true;
                }

                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Adjustment ??=
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustment();
                ceComp.AdjustedFanil.Adjustment.MainEntityPEandFte.Add(item);
            }
        }

        // ─── 3.2.4.1(c): 국가간 손익 조정 → CrossBorderAdjustments ─────
        // 헤더행: "(c) 국가간 손익 조정", 컬럼헤더 +1행, 데이터 +2행~
        // C(3)=Basis, F(6)=OtherTIN, I(9)=ResCountryCode, L(12)=Additions, P(16)=Reductions
        private void Map3241c(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "국가간 손익 조정");
            if (rHdr < 0)
                return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 15, lastRow); r++)
            {
                var basisRaw = ws.Cell(r, 3).GetString()?.Trim(); // C: 조정근거
                var otherTinRaw = ws.Cell(r, 6).GetString()?.Trim(); // F: 상대TIN
                var adds = ws.Cell(r, 12).GetString()?.Trim(); // L: 가산액
                var reds = ws.Cell(r, 16).GetString()?.Trim(); // P: 차감액

                if (
                    string.IsNullOrEmpty(basisRaw)
                    && string.IsNullOrEmpty(otherTinRaw)
                    && string.IsNullOrEmpty(adds)
                    && string.IsNullOrEmpty(reds)
                )
                    break;
                if (string.IsNullOrEmpty(basisRaw))
                    continue;

                if (!TryParseEnum<Globe.CrossBorderAdjustmentsEnumType>(basisRaw, out var basis))
                {
                    errors.Add(
                        $"[{fileName}] 3.2.4.1(c) 조정근거 '{basisRaw}' 파싱 실패 — GIR1801~GIR1802 중 하나여야 함"
                    );
                    continue;
                }

                var item =
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentCrossBorderAdjustments
                    {
                        Basis = basis,
                        OtherTin = ParseTin(otherTinRaw ?? ""),
                        Additions = HasNonZeroValue(adds) ? adds : null,
                        Reductions = HasNonZeroValue(reds) ? reds : null,
                    };

                var ccRaw = ws.Cell(r, 9).GetString()?.Trim(); // I: 소재지국 (I4)
                if (
                    !string.IsNullOrEmpty(ccRaw)
                    && TryParseEnum<Globe.CountryCodeType>(ccRaw, out var cc)
                )
                {
                    item.ResCountryCode = cc;
                    item.ResCountryCodeSpecified = true;
                }

                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Adjustment ??=
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustment();
                ceComp.AdjustedFanil.Adjustment.CrossBorderAdjustments.Add(item);
            }
        }

        // ─── 3.2.4.1(d): 최종모기업 글로벌최저한세소득 감액 → UpeAdjustments ──
        // 헤더행: "최종모기업의 글로벌최저한세소득 감액", 컬럼헤더 +1행, 데이터 +2행~ (행추가 가능)
        // E(5)=Basis, H(8)=소유지분보유자TIN, K(11)=직접보유비율(%), N(14)=소득차감액
        private void Map3241d(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rHdr = FindRow(ws, "최종모기업의 글로벌최저한세소득 감액");
            if (rHdr < 0)
                return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? rHdr + 20;
            for (int r = rHdr + 2; r <= System.Math.Min(rHdr + 20, lastRow); r++)
            {
                var basisRaw = ws.Cell(r, 5).GetString()?.Trim(); // E: 감액근거
                var ownerRaw = ws.Cell(r, 8).GetString()?.Trim(); // H: 소유지분보유자TIN
                var pctRaw = ws.Cell(r, 11).GetString()?.Trim(); // K: 직접보유비율(%)
                var reductRaw = ws.Cell(r, 14).GetString()?.Trim(); // N: 소득차감액

                if (
                    string.IsNullOrEmpty(basisRaw)
                    && string.IsNullOrEmpty(ownerRaw)
                    && string.IsNullOrEmpty(pctRaw)
                    && string.IsNullOrEmpty(reductRaw)
                )
                    break;
                if (string.IsNullOrEmpty(basisRaw))
                    continue;

                if (!TryParseEnum<Globe.UpeAdjustmentsBasisEnumType>(basisRaw, out var basis))
                {
                    errors.Add(
                        $"[{fileName}] 3.2.4.1(d) 감액근거 '{basisRaw}' 파싱 실패 — GIR1901 등 유효한 코드여야 함"
                    );
                    continue;
                }

                var upe =
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustments
                    {
                        Basis = basis,
                    };

                // N열: 소득차감액 → Reductions.Amount
                if (!string.IsNullOrEmpty(reductRaw))
                    upe.Reductions =
                        new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsReductions
                        {
                            Amount = reductRaw,
                        };

                // H열 + K열: 소유지분보유자 → IdentificationOfOwners
                if (!string.IsNullOrEmpty(ownerRaw) || !string.IsNullOrEmpty(pctRaw))
                {
                    var owner = ParseOwner(ownerRaw, pctRaw, errors, fileName);
                    if (owner != null)
                        upe.IdentificationOfOwners.Add(owner);
                }

                ceComp.AdjustedFanil ??= new Globe.EtrComputationTypeCeComputationAdjustedFanil();
                ceComp.AdjustedFanil.Adjustment ??=
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustment();
                ceComp.AdjustedFanil.Adjustment.UpeAdjustments.Add(upe);
            }
        }

        // ─── 소유지분보유자 파싱 ────────────────────────────────────────────
        // 개인: 개인,주주수,거주지국[,세율]
        // 법인: 법인,TIN값,유형,발급국가[,세율[,ExTypeOfEntity코드]]
        //       세율 생략 시 빈 칸 허용: 법인,TIN,유형,발급국가,,ExType
        private static Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwners ParseOwner(
            string ownerRaw,
            string pctRaw,
            List<string> errors,
            string fileName
        )
        {
            var owner =
                new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwners();

            // K열: 직접보유비율(%)
            var pct = ParsePercentage(pctRaw);
            if (pct.HasValue)
                owner.OwnershipPercentage = pct.Value;

            if (string.IsNullOrEmpty(ownerRaw))
                return owner;

            var parts = ownerRaw.Split(',');
            var kind = parts[0].Trim();

            if (kind == "개인")
            {
                // 개인,주주수,거주지국[,세율]
                var ind =
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwnersIndOwners
                    {
                        NumOfOwners = parts.Length > 1 ? parts[1].Trim() : "1",
                    };
                if (
                    parts.Length > 2
                    && !string.IsNullOrEmpty(parts[2].Trim())
                    && TryParseEnum<Globe.CountryCodeType>(parts[2].Trim(), out var cc)
                )
                {
                    ind.ResCountryCode = cc;
                    ind.ResCountryCodeSpecified = true;
                }
                if (parts.Length > 3)
                {
                    var rate = ParsePercentage(parts[3]);
                    if (rate.HasValue)
                    {
                        ind.TaxRate = rate.Value;
                        ind.TaxRateSpecified = true;
                    }
                }
                owner.IndOwners = ind;
            }
            else if (kind == "법인")
            {
                // 법인,TIN값,유형,발급국가[,세율[,ExTypeOfEntity코드]]
                if (parts.Length < 4)
                {
                    errors.Add(
                        $"[{fileName}] 3.2.4.1(d) 법인 소유자 형식 오류 '{ownerRaw}' — 법인,TIN값,유형,발급국가[,세율[,ExType]] 형식이어야 함"
                    );
                    return owner;
                }
                var tinRaw = $"{parts[1].Trim()},{parts[2].Trim()},{parts[3].Trim()}";
                var tin = ParseTin(tinRaw);
                var entity =
                    new Globe.EtrComputationTypeCeComputationAdjustedFanilAdjustmentUpeAdjustmentsIdentificationOfOwnersEntityOwner
                    {
                        Tin = tin,
                        ResCountryCode = tin.IssuedBySpecified ? tin.IssuedBy : default,
                    };
                // parts[4]: 세율 (빈 문자열이면 스킵)
                if (parts.Length > 4)
                {
                    var rate = ParsePercentage(parts[4]);
                    if (rate.HasValue)
                    {
                        entity.TaxRate = rate.Value;
                        entity.TaxRateSpecified = true;
                    }
                }
                // parts[5]: ExTypeOfEntity (세율 스킵해도 파싱)
                if (
                    parts.Length > 5
                    && !string.IsNullOrEmpty(parts[5].Trim())
                    && TryParseEnum<Globe.ExTypeOfEntityEnumType>(parts[5].Trim(), out var exType)
                )
                {
                    entity.ExTypeOfEntity = exType;
                    entity.ExTypeOfEntitySpecified = true;
                }
                owner.EntityOwner = entity;
            }
            else
            {
                errors.Add(
                    $"[{fileName}] 3.2.4.1(d) 소유자 유형 '{kind}' 오류 — '개인' 또는 '법인'으로 시작해야 함"
                );
            }

            return owner;
        }

        // ─── 3.2.4.3: 구성기업별 선택 ────────────────────────────────────
        // 매년선택 (a,b,c): "2. 매년 선택" 행 기준 오프셋, O(15) = 여/부
        // 5년선택 (d~i):   "3. 5년 선택" 행 기준 오프셋, O(15)=선택연도, Q(17)=취소연도
        // 기타선택 (j):    "6. 기타선택" 행 기준 오프셋, O(15)=선택연도, Q(17)=취소연도
        // 공정가액조정 (k): "k. 공정가액조정" 행 이후 데이터 행, E(5)=사업연도, J(10)=(i)/(ii)
        private void Map3243(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.3");
            if (rSec < 0)
                return;

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            var el = ceComp.Elections;

            // ── 매년 선택 (a, b, c) ─────────────────────────────────────────
            var rMaeYear = FindRow(ws, "2. 매년 선택", rSec);
            if (rMaeYear >= 0)
            {
                // a: 간소화 계산 → SimplCalculations
                var vA = ws.Cell(rMaeYear, 15).GetString()?.Trim(); // O120
                if (!string.IsNullOrEmpty(vA))
                {
                    el.SimplCalculations = ParseBool(vA);
                    el.SimplCalculationsSpecified = true;
                }
                // b: 채무면제이익 제외 → Art321
                var vB = ws.Cell(rMaeYear + 1, 15).GetString()?.Trim(); // O121
                if (!string.IsNullOrEmpty(vB))
                {
                    el.Art321 = ParseBool(vB);
                    el.Art321Specified = true;
                }
                // c: 이연법인세비용 미반영(매년) → KArt447C
                var vC = ws.Cell(rMaeYear + 2, 15).GetString()?.Trim(); // O122
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
                            RevocationYearSpecified = ry.HasValue,
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
                            RevocationYearSpecified = ry.HasValue,
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
                            RevocationYearSpecified = ry.HasValue,
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
                            RevocationYearSpecified = ry.HasValue,
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
                        RevocationYearSpecified = ry.HasValue,
                    };
            }

            // ── k: 공정가액조정 선택 → Art634[] ───────────────────────────────
            // B열: "k. 공정가액조정", 이후 +2행부터 데이터 (헤더행 1개 건너뜀)
            // E(5)=FYTriggerEvent(날짜) [병합 E:H의 앵커]
            // I(9)=(i)/(ii) [병합 I:R의 앵커 — 이전 J(10) 잘못 기재 수정]
            // Inclusion이 실제 (i)/(ii) 파싱 성공한 경우만 Art634 엔트리 추가 (빈 Inclusion 방지)
            var rK = FindRow(ws, "k. 공정가액조정", rSec);
            if (rK >= 0)
            {
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? rK + 10;
                for (int r = rK + 2; r <= System.Math.Min(rK + 10, lastRow); r++)
                {
                    var eyRaw = ws.Cell(r, 5).GetString()?.Trim(); // E: 사업연도
                    var inclRaw = ws.Cell(r, 9).GetString()?.Trim(); // I: (i)/(ii) 병합 앵커
                    if (string.IsNullOrEmpty(eyRaw) && string.IsNullOrEmpty(inclRaw))
                        break;
                    if (!TryParseDate(eyRaw, out var fy))
                        continue;

                    var incl = new Globe.EtrComputationTypeCeComputationElectionsArt634Inclusion();
                    bool hasIncl = false;
                    if (inclRaw == "(i)" || inclRaw == "i")
                    {
                        incl.Art634CI = true;
                        incl.Art634CISpecified = true;
                        hasIncl = true;
                    }
                    else if (inclRaw == "(ii)" || inclRaw == "ii")
                    {
                        incl.Art634CIi = true;
                        incl.Art634CIiSpecified = true;
                        hasIncl = true;
                    }

                    el.Art634.Add(
                        new Globe.EtrComputationTypeCeComputationElectionsArt634
                        {
                            FyTriggerEvent = fy,
                            Inclusion = hasIncl ? incl : null,
                        }
                    );
                }
            }
        }

        /// <summary>
        /// MapOneEntity 끝에서 호출: ceComp의 하위 객체 중 실제로 데이터가 없는 것들을
        /// null로 되돌려 빈 태그(&lt;Elections /&gt; 등)가 XML에 나오지 않도록 정리.
        /// </summary>
        private static void CleanupEmptyCeComp(Globe.EtrComputationTypeCeComputation ceComp)
        {
            if (ceComp.Elections is { } e0)
            {
                bool hasAny =
                    e0.SimplCalculationsSpecified
                    || e0.Art321Specified
                    || e0.KArt447CSpecified
                    || e0.Art153 != null
                    || e0.Art321B != null
                    || e0.Art321C != null
                    || e0.Art634.Count > 0
                    || e0.AggregatedReporting != null
                    || e0.Art447 != null
                    || e0.Art456 != null
                    || e0.Art75.Count > 0
                    || e0.Art76.Count > 0;
                if (!hasAny)
                    ceComp.Elections = null;
            }
        }

        // ─── 3.2.4.4: 국제해운소득·결손 제외 → CEComputation.NetGlobeIncome.IntShippingIncome ──
        // rSec+3: 항목2 국제해운 Category(O=15, 쉼표구분), rSec+4: Revenue, rSec+5: Costs, rSec+6: Total
        // rSec+7: 항목6 부수소득 Category(O=15 단일), rSec+8: Revenue, rSec+9: Costs, rSec+10: Total
        // rSec+11: 인건비(PayrollCosts), rSec+12: 유형자산(TangibleAssets), rSec+13: CoveredTaxes
        private void Map3244(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.4");
            if (rSec < 0)
                return;

            string V(int offset) => ws.Cell(rSec + offset, 15).GetString()?.Trim();

            // InternationalShipIncome
            var intCatRaw = V(3);
            var intRevenue = V(4);
            var intCosts = V(5);
            var intTotal = V(6);

            // QualifiedAncShipIncome
            var ancCatRaw = V(7);
            var ancRevenue = V(8);
            var ancCosts = V(9);
            var ancTotal = V(10);

            // SubstanceExclusion
            var payroll = V(11);
            var tangible = V(12);

            // CoveredTaxes
            var covTaxes = V(13);

            // 모두 빈 값이면 스킵
            if (
                new[]
                {
                    intCatRaw,
                    intRevenue,
                    intCosts,
                    intTotal,
                    ancCatRaw,
                    ancRevenue,
                    ancCosts,
                    ancTotal,
                    payroll,
                    tangible,
                    covTaxes,
                }.All(string.IsNullOrEmpty)
            )
                return;

            ceComp.NetGlobeIncome ??= new Globe.EtrComputationTypeCeComputationNetGlobeIncome();
            var ngi = ceComp.NetGlobeIncome;

            var shipping =
                new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncome
                {
                    CoveredTaxes = NullIfEmpty(covTaxes),
                };

            // InternationalShipIncome — 실질 소득 데이터(Total/Revenue/Costs)가 있을 때만 생성
            // XSD상 Total/Category/Revenue/Costs 모두 required이므로 일부만 있으면 안 됨.
            // Category만 있고 income 없으면 사용자 입력 미완성으로 간주, 생략.
            bool hasIntIncome = HasNonZeroValue(intTotal) || HasNonZeroValue(intRevenue) || HasNonZeroValue(intCosts);
            if (hasIntIncome)
            {
                shipping.InternationalShipIncome =
                    new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncomeInternationalShipIncome
                    {
                        Total = NullIfEmpty(intTotal),
                        Revenue = NullIfEmpty(intRevenue),
                        Costs = NullIfEmpty(intCosts),
                    };
            }
            else if (!string.IsNullOrEmpty(intCatRaw))
            {
                errors.Add(
                    $"[{fileName}] 3.2.4.4 국제해운소득 유형('{intCatRaw}')은 입력됐으나 Total/Revenue/Costs가 없어 생략됨"
                );
            }

            // QualifiedAncShipIncome — 동일 규칙
            bool hasAncIncome = HasNonZeroValue(ancTotal) || HasNonZeroValue(ancRevenue) || HasNonZeroValue(ancCosts);
            if (hasAncIncome)
            {
                shipping.QualifiedAncShipIncome =
                    new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncomeQualifiedAncShipIncome
                    {
                        Total = NullIfEmpty(ancTotal),
                        Revenue = NullIfEmpty(ancRevenue),
                        Costs = NullIfEmpty(ancCosts),
                    };
            }
            else if (!string.IsNullOrEmpty(ancCatRaw))
            {
                errors.Add(
                    $"[{fileName}] 3.2.4.4 적격국제해운부수소득 유형('{ancCatRaw}')은 입력됐으나 Total/Revenue/Costs가 없어 생략됨"
                );
            }

            // SubstanceExclusion — 0이 아닌 의미있는 값이 있을 때만 생성
            if (HasNonZeroValue(payroll) || HasNonZeroValue(tangible))
            {
                shipping.SubstanceExclusion =
                    new Globe.EtrComputationTypeCeComputationNetGlobeIncomeIntShippingIncomeSubstanceExclusion
                    {
                        PayrollCosts = NullIfEmpty(payroll),
                        TangibleAssets = NullIfEmpty(tangible),
                    };
            }

            // 모든 부분구조가 비어있으면 IntShippingIncome 자체 생략
            if (
                shipping.InternationalShipIncome == null
                && shipping.QualifiedAncShipIncome == null
                && shipping.SubstanceExclusion == null
                && string.IsNullOrEmpty(shipping.CoveredTaxes)
            )
                return;

            // CoveredTaxes는 XSD required — 빈 값이면 "0" 기본값
            if (string.IsNullOrEmpty(shipping.CoveredTaxes))
                shipping.CoveredTaxes = "0";

            ngi.IntShippingIncome = shipping;

            // 국제해운소득 유형(Category) — InternationalShipIncome 생성된 경우만 추가
            if (!string.IsNullOrEmpty(intCatRaw) && shipping.InternationalShipIncome != null)
            {
                var intCodes = System.Text.RegularExpressions.Regex.Matches(intCatRaw, @"GIR\d+");
                if (intCodes.Count == 0)
                    errors.Add(
                        $"[{fileName}] 3.2.4.4 국제해운소득 유형 파싱 실패: '{intCatRaw}' (GIR2101~GIR2106)"
                    );
                foreach (System.Text.RegularExpressions.Match m in intCodes)
                {
                    if (TryParseEnum<Globe.IntShipCategoryEnumType>(m.Value, out var cat))
                        shipping.InternationalShipIncome.Category.Add(cat);
                    else
                        errors.Add(
                            $"[{fileName}] 3.2.4.4 국제해운소득 유형 파싱 실패: '{m.Value}' (GIR2101~GIR2106)"
                        );
                }
            }

            // 적격국제해운부수소득 유형(Category) — QualifiedAncShipIncome 생성된 경우만 설정
            if (!string.IsNullOrEmpty(ancCatRaw) && shipping.QualifiedAncShipIncome != null)
            {
                if (TryParseEnum<Globe.AncShipCategoryEnumType>(ancCatRaw, out var ancCat))
                {
                    shipping.QualifiedAncShipIncome.Category = ancCat;
                }
                else
                    errors.Add(
                        $"[{fileName}] 3.2.4.4 적격국제해운부수소득 유형 파싱 실패: '{ancCatRaw}' (GIR2201~GIR2205)"
                    );
            }
        }

        // ─── 3.2.4.5: 과세분배방법 적용 선택 관련 정보 → CEComputation.Elections.Art76[] ──
        // 헤더: rSec+2(R156), 데이터 5행: rSec+3~rSec+7
        // B(2)=주주구성기업TIN, E(5)=투자기업TIN, G(7)=ActualDeemedDist,
        // K(11)=LocalCreditableTaxGross, N(14)=ShareOfUndistNetGlobeInc
        // Status/ElectionYear/RevocationYear: 3.2.4.3 h행(r5Year+5)에서 읽음
        private void Map3245(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.5");
            if (rSec < 0)
                return;

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
                    if (st.HasValue)
                        art76Status = st.Value;
                    if (ey.HasValue)
                        art76Ey = ey.Value;
                    art76Ry = ry;
                }
            }

            ceComp.Elections ??= new Globe.EtrComputationTypeCeComputationElections();
            var el = ceComp.Elections;

            // 데이터 5행: rSec+3 ~ rSec+7
            for (int offset = 3; offset <= 7; offset++)
            {
                var r = rSec + offset;
                var eRaw = ws.Cell(r, 5).GetString()?.Trim(); // E: 투자기업 TIN (E4)
                var gRaw = ws.Cell(r, 7).GetString()?.Trim(); // G: ActualDeemedDist (G4)
                var kRaw = ws.Cell(r, 11).GetString()?.Trim(); // K: LocalCreditableTaxGross (K4)
                var nRaw = ws.Cell(r, 14).GetString()?.Trim(); // N: ShareOfUndistNetGlobeInc (N4)

                if (string.IsNullOrEmpty(eRaw) && string.IsNullOrEmpty(gRaw))
                    continue;

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
                    var share = ParsePercentage(nRaw);
                    if (share.HasValue)
                        a76.ShareOfUndistNetGlobeInc = share.Value;
                    else
                        errors.Add($"[{fileName}] 3.2.4.5 지분율 형식 오류: '{nRaw}' (0~1 사이 decimal 입력, 예: 0.5)");
                }

                if (!string.IsNullOrEmpty(eRaw))
                {
                    // ParseTin은 항상 비-null 반환 (빈 입력은 NoTin)
                    a76.InvestmentEntityTin = ParseTin(eRaw);
                }

                el.Art76.Add(a76);
            }
        }

        // ─── 3.2.4.6: 그 밖의 회계기준 → CEComputation.OtherFas ──────────────
        // 헤더: rSec+1(R163), 데이터 행: rSec+2 이후 (B가 빈 행까지)
        // B(2)=CE TIN, K(11)=OtherFas (적용 인정회계기준/공인회계기준)
        // ceComp.Tin.Value와 일치하는 행의 K값을 사용, 없으면 첫 번째 비어있지 않은 K값
        private void Map3246(
            IXLWorksheet ws,
            Globe.EtrComputationTypeCeComputation ceComp,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.2.4.6");
            if (rSec < 0)
                return;

            var currentTin = ceComp.Tin?.Value?.Trim();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 300;
            string matched = null;
            string firstFound = null;

            for (int r = rSec + 2; r <= lastRow; r++)
            {
                var bRaw = ws.Cell(r, 2).GetString()?.Trim();
                var kRaw = ws.Cell(r, 11).GetString()?.Trim();

                if (string.IsNullOrEmpty(bRaw) && string.IsNullOrEmpty(kRaw))
                    break;
                if (string.IsNullOrEmpty(kRaw))
                    continue;

                if (firstFound == null)
                    firstFound = kRaw;

                if (
                    !string.IsNullOrEmpty(currentTin)
                    && !string.IsNullOrEmpty(bRaw)
                    && bRaw.Contains(currentTin)
                )
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
        // 1. a,b,c → O(15)열 / 2+3 모기업 정보 → 통합 셀(O7, 2번 섹션 앵커)
        // 통합 셀 포맷(모기업 1건, ',' 구분 9필드, 주주 구분 ';'):
        //   TIN값, TIN유형(GIR300x), 발급국가(ISO2), 모기업소재지국(ISO2),
        //   OtherOwnershipAllocation, InclusionRatio, TopUpTaxShare, IirOffSet, TopUpTax
        private void Map341(
            IXLWorksheet ws,
            Globe.JurisdictionSectionType js,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.4.1");
            if (rSec < 0)
                return;

            // 1. 추가세액 배분 그룹기업 (LTCE) a,b,c → O(15)열
            var rLtce = FindRow(ws, "1. 추가세액 배분", rSec);
            if (rLtce < 0)
                return;

            var tinRaw = ws.Cell(rLtce, 15).GetString()?.Trim(); // a: TIN (O171)
            var ngiRaw = ws.Cell(rLtce + 1, 15).GetString()?.Trim(); // b: NetGlobeIncome (O172)
            var tutRaw = ws.Cell(rLtce + 2, 15).GetString()?.Trim(); // c: TopUpTax (O173)
            if (string.IsNullOrEmpty(tinRaw))
                return;

            // IIR 엔트리 생성
            var iir = new Globe.LowTaxJurisdictionTypeLtceIir
            {
                NetGlobeIncome = string.IsNullOrEmpty(ngiRaw) ? null : ngiRaw,
                TopUpTax = string.IsNullOrEmpty(tutRaw) ? "0" : tutRaw,
            };

            // 2+3. 모기업 정보 통합 셀 → ParentEntity[]
            // "2. 적격소득산입규칙을 적용해야 하는 모기업" 행 앵커 O열
            var rParent = FindRow(ws, "2. 적격소득산입규칙", rSec);
            if (rParent < 0)
                rParent = FindRow(ws, "적격소득산입규칙", rSec);
            if (rParent > 0)
            {
                var parentRaw = ws.Cell(rParent, 15).GetString()?.Trim(); // O174
                if (!string.IsNullOrEmpty(parentRaw))
                    ParseParentEntities(parentRaw, iir, rParent, errors, fileName);
            }

            var ltce = new Globe.LowTaxJurisdictionTypeLtce { Tin = ParseTin(tinRaw) };
            ltce.Iir.Add(iir);

            js.LowTaxJurisdiction ??= new Globe.LowTaxJurisdictionType { TopUpTaxAmount = "0" };
            js.LowTaxJurisdiction.Ltce.Add(ltce);
        }

        /// <summary>
        /// 3.4.1 모기업 통합 셀 파싱.
        /// 필드 9개(',' 구분): TIN값, TIN유형, 발급국가, 소재지국,
        ///   OtherOwnershipAllocation, InclusionRatio, TopUpTaxShare, IirOffSet, TopUpTax
        /// 여러 모기업은 ';'로 구분.
        /// </summary>
        private static void ParseParentEntities(
            string cellValue,
            Globe.LowTaxJurisdictionTypeLtceIir iir,
            int row,
            List<string> errors,
            string fileName
        )
        {
            var shareholders = cellValue.Split(
                ';',
                System.StringSplitOptions.RemoveEmptyEntries | System.StringSplitOptions.TrimEntries
            );

            for (int i = 0; i < shareholders.Length; i++)
            {
                var parts = shareholders[i].Split(',', System.StringSplitOptions.TrimEntries);
                if (parts.Length == 0 || parts.All(string.IsNullOrEmpty))
                    continue;

                var tinVal = parts.Length >= 1 ? parts[0] : null;
                var tinType = parts.Length >= 2 ? parts[1] : null;
                var issuedBy = parts.Length >= 3 ? parts[2] : null;
                var resCc = parts.Length >= 4 ? parts[3] : null;
                var other = parts.Length >= 5 ? parts[4] : null;
                var incRatio = parts.Length >= 6 ? parts[5] : null;
                var tutShare = parts.Length >= 7 ? parts[6] : null;
                var iirOff = parts.Length >= 8 ? parts[7] : null;
                var topUpTax = parts.Length >= 9 ? parts[8] : null;

                if (string.IsNullOrEmpty(tinVal))
                    continue;

                // TIN 조립
                Globe.TinType tin;
                if (tinVal.Equals("NOTIN", System.StringComparison.OrdinalIgnoreCase))
                {
                    tin = NoTin();
                }
                else
                {
                    tin = new Globe.TinType { Value = tinVal };
                    if (
                        !string.IsNullOrEmpty(tinType)
                        && TryParseEnum<Globe.TinEnumType>(tinType, out var tinEnum)
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
                }

                // 소재지국
                Globe.CountryCodeType cc = default;
                if (!string.IsNullOrEmpty(resCc) && !System.Enum.TryParse(resCc, true, out cc))
                {
                    errors.Add(
                        $"[{fileName}] [3.4.1 모기업{i + 1}] 소재지국 '{resCc}' 파싱 실패 (O{row})"
                    );
                }

                // 소득산입비율 (decimal)
                var inclusionRatio = ParsePercentage(incRatio) ?? 0m;

                iir.ParentEntity.Add(
                    new Globe.LowTaxJurisdictionTypeLtceIirParentEntity
                    {
                        Tin = tin,
                        ResCountryCode = cc,
                        OtherOwnershipAllocation = string.IsNullOrEmpty(other) ? "0" : other,
                        InclusionRatio = inclusionRatio,
                        TopUpTaxShare = string.IsNullOrEmpty(tutShare) ? "0" : tutShare,
                        IirOffSet = string.IsNullOrEmpty(iirOff) ? "0" : iirOff,
                        TopUpTax = string.IsNullOrEmpty(topUpTax) ? "0" : topUpTax,
                    }
                );
            }
        }

        // ─── 3.4.2: 소득산입보완규칙(UTPR) → LowTaxJurisdiction.Utpr ─────────
        // 1: 납세자번호(XML 미포함), 2: Article2.5.1TopUpTax, 3: TotalUTPRTopUpTax — 모두 O(15)열
        private void Map342(
            IXLWorksheet ws,
            Globe.JurisdictionSectionType js,
            List<string> errors,
            string fileName
        )
        {
            var rSec = FindRow(ws, "3.4.2");
            if (rSec < 0)
                return;

            // 3.4.2.2: "제73조제3항제2호에 따라" → Article251TopUpTax
            var r2 = FindRow(ws, "3항제2호", rSec);
            if (r2 < 0)
                r2 = FindRow(ws, "납부액을 차감", rSec); // 폴백
            // 3.4.2.3: 소득산입보완규칙 추가세액 합계 → TotalUtprTopUpTax
            var r3 = FindRow(ws, "소득산입보완규칙 추가세액 합계", rSec);

            var art251 = r2 > 0 ? ws.Cell(r2, 15).GetString()?.Trim() : null; // O184
            var totalUtpr = r3 > 0 ? ws.Cell(r3, 15).GetString()?.Trim() : null; // O185
            if (string.IsNullOrEmpty(art251) && string.IsNullOrEmpty(totalUtpr))
                return;

            js.LowTaxJurisdiction ??= new Globe.LowTaxJurisdictionType { TopUpTaxAmount = "0" };
            js.LowTaxJurisdiction.Utpr ??= new Globe.LowTaxJurisdictionTypeUtpr();
            js.LowTaxJurisdiction.Utpr.UtprCalculation =
                new Globe.LowTaxJurisdictionTypeUtprUtprCalculation
                {
                    Article251TopUpTax = art251 ?? "0",
                    TotalUtprTopUpTax = totalUtpr ?? "0",
                };
        }

        // 선택 사업연도(O=15)/취소 사업연도(Q=17) 읽기 헬퍼
        // Status: 선택연도가 있으면 true, 없고 취소연도만 있으면 false
        private static (bool? status, System.DateTime? ey, System.DateTime? ry) ReadYearElection(
            IXLWorksheet ws,
            int row
        )
        {
            var eyRaw = ws.Cell(row, 15).GetString()?.Trim(); // O
            var ryRaw = ws.Cell(row, 17).GetString()?.Trim(); // Q
            if (string.IsNullOrEmpty(eyRaw) && string.IsNullOrEmpty(ryRaw))
                return (null, null, null);

            TryParseDate(eyRaw, out var ey);
            TryParseDate(ryRaw, out var ry);
            bool status = !string.IsNullOrEmpty(eyRaw);
            return (
                status,
                ey == default ? (System.DateTime?)null : ey,
                ry == default ? (System.DateTime?)null : ry
            );
        }

        // Art153 전용
        private static Globe.EtrComputationTypeCeComputationElectionsArt153 ReadYearElectionArt153(
            IXLWorksheet ws,
            int row
        )
        {
            var (status, ey, ry) = ReadYearElection(ws, row);
            if (!status.HasValue)
                return null;
            return new Globe.EtrComputationTypeCeComputationElectionsArt153
            {
                Status = status.Value,
                ElectionYear = ey ?? default,
                ElectionYearSpecified = ey.HasValue,
                RevocationYear = ry ?? default,
                RevocationYearSpecified = ry.HasValue,
            };
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
                var v = ws.Cell(r, 2).GetString();
                if (v != null && v.Contains(contains))
                    return r;
            }
            return -1;
        }
    }
}
