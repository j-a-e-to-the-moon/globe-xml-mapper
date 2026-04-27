using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 1.3.3 기업구조 변동 — CorporateStructureTypeCeOwnershipChange 매핑.
    /// 컬럼 헤더("1. 구성기업") 다음 행부터 데이터로 iteration.
    /// 각 행을 CE.OwnershipChange에 추가 (CE는 1.3.2.1에서 이미 생성된 것 참조).
    /// </summary>
    public class Mapping_1_3_3 : MappingBase
    {
        private const string COLUMN_HEADER_ANCHOR = "1. 구성기업"; // B열 컬럼헤더

        public Mapping_1_3_3()
            : base(null) { }

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??=
                new Globe.CorporateStructureType();
            var cs = globe.GlobeBody.GeneralSection.CorporateStructure;

            var headerRow = FindAnchorRow(ws, COLUMN_HEADER_ANCHOR);
            if (headerRow < 0) return;
            var dataStartRow = headerRow + 1;
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? dataStartRow;

            for (int row = dataStartRow; row <= lastRow; row++)
            {
                // B: 구성기업 상호 (표시용, 매칭은 D열 TIN 기준)
                var name = ws.Cell(row, 2).GetString()?.Trim();
                // D: 납세자번호 — CE 매칭 기준 ("번호,유형,국가" 형식)
                var tinRaw = ws.Cell(row, 4).GetString()?.Trim();
                // F: 변동효력발생일
                var dateRaw = ws.Cell(row, 6).GetString()?.Trim();
                // H: 변동 전 기업유형 (multi, IdGlobeStatusEnumType)
                var preTypeRaw = ws.Cell(row, 8).GetString()?.Trim();
                // J: 소유지분 보유 기업 TIN ("번호,유형,국가" 형식)
                var ownerTinRaw = ws.Cell(row, 10).GetString()?.Trim();
                // M: 변동 전 소유지분(%)
                var prePctRaw = ws.Cell(row, 13).GetString()?.Trim();
                // P: 소유지분 보유 기업 유형 (OwnershipTypeEnumType)
                var ownerTypeRaw = ws.Cell(row, 16).GetString()?.Trim();

                if (string.IsNullOrEmpty(tinRaw))
                    continue;

                // ── CE 찾기 (2번 납세자번호 기준) ────────────────────────
                var ce = FindCeByTin(cs, tinRaw);
                if (ce == null)
                {
                    errors.Add(
                        $"[{fileName}] 행{row}: 1.3.2.1에 TIN '{tinRaw.Split(',')[0].Trim()}'인 CE 없음 — '{name}' 건너뜀"
                    );
                    continue;
                }

                // ── OwnershipChange 생성 ──────────────────────────────────
                var change = new Globe.CorporateStructureTypeCeOwnershipChange();

                if (!string.IsNullOrEmpty(dateRaw))
                {
                    if (TryParseDate(dateRaw, out var changeDate))
                        change.ChangeDate = changeDate;
                    else
                        errors.Add(
                            $"[{fileName}] 행{row}: 변동효력발생일 날짜 형식 오류 '{dateRaw}' (예: 2024-03-31)"
                        );
                }

                if (!string.IsNullOrEmpty(preTypeRaw))
                {
                    foreach (
                        var code in preTypeRaw.Split(
                            ',',
                            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
                        )
                    )
                    {
                        if (TryParseEnum<Globe.IdGlobeStatusEnumType>(code, out var status))
                            change.PreGlobeStatus.Add(status);
                        else
                            errors.Add($"[{fileName}] 행{row} PreGlobeStatus 변환 실패: '{code}'");
                    }
                }

                if (!string.IsNullOrEmpty(ownerTinRaw))
                {
                    var preOwn = new Globe.CorporateStructureTypeCeOwnershipChangePreOwnership
                    {
                        Tin = ParseTin(ownerTinRaw),
                    };

                    if (!string.IsNullOrEmpty(ownerTypeRaw))
                    {
                        if (
                            TryParseEnum<Globe.OwnershipTypeEnumType>(
                                ownerTypeRaw,
                                out var ownerType
                            )
                        )
                            preOwn.OwnershipType = ownerType;
                        else
                            errors.Add(
                                $"[{fileName}] 행{row} OwnershipType 변환 실패: '{ownerTypeRaw}'"
                            );
                    }

                    var prePct = ParsePercentage(prePctRaw);
                    if (prePct.HasValue)
                        preOwn.PreOwnershipPercentage = prePct.Value;

                    change.PreOwnership.Add(preOwn);
                }

                ce.OwnershipChange.Add(change);
            }
        }

        private static Globe.CorporateStructureTypeCe FindCeByTin(
            Globe.CorporateStructureType cs,
            string tinRaw
        )
        {
            var tinValue = tinRaw.Split(',')[0].Trim();
            return cs.Ce.FirstOrDefault(c => c.Id?.Tin.Any(t => t.Value == tinValue) == true);
        }

        // B열에서 anchor 텍스트를 포함하는 첫 행 반환 (-1 = 없음).
        private static int FindAnchorRow(IXLWorksheet ws, string contains)
        {
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 200;
            for (int r = 1; r <= lastRow; r++)
            {
                var v = ws.Cell(r, 2).GetString() ?? "";
                if (v.Contains(contains)) return r;
            }
            return -1;
        }
    }
}
