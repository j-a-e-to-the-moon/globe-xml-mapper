using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// "구성기업" 시트 리더.
    /// 헤더(3행): D=기업 TIN / E=국가명 / F=하위그룹 유형 / H=하위그룹 최상위기업 TIN
    ///
    /// 구성기업(entity) TIN으로 합산단위(group) 식별자를 조회:
    ///   entityTin → (country, subGroupTypes[], subGroupTin)
    ///
    /// 합산단위 자체(그룹 블록 → JurisdictionSection)는 (country, subGroupTypes, subGroupTin) 조합으로 식별.
    /// </summary>
    public sealed class EntityGroupMap
    {
        public sealed class Entry
        {
            public string EntityTin { get; init; }
            public Globe.CountryCodeType? Country { get; init; }
            public string SubGroupTypesRaw { get; init; } // 콤마 구분된 유형 원본
            public string SubGroupTin { get; init; } // 하위그룹 최상위기업 TIN (없을 수 있음)
        }

        private readonly Dictionary<string, Entry> _byEntityTin = new();

        public static EntityGroupMap Load(IXLWorkbook workbook, List<string> errors = null)
        {
            var map = new EntityGroupMap();
            if (!workbook.TryGetWorksheet("구성기업", out var ws))
                return map;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 3;
            for (int r = 4; r <= lastRow; r++)
            {
                var tin = ws.Cell(r, 4).GetString()?.Trim();   // D열
                var country = ws.Cell(r, 5).GetString()?.Trim(); // E열
                var sgType = ws.Cell(r, 6).GetString()?.Trim(); // F열
                var sgTin = ws.Cell(r, 8).GetString()?.Trim();  // H열

                if (string.IsNullOrEmpty(tin))
                    continue;

                Globe.CountryCodeType? countryCode = null;
                if (
                    !string.IsNullOrEmpty(country)
                    && Enum.TryParse<Globe.CountryCodeType>(country, true, out var parsed)
                )
                {
                    countryCode = parsed;
                }
                else if (!string.IsNullOrEmpty(country))
                {
                    errors?.Add($"[구성기업 R{r}] 국가코드 '{country}' 파싱 실패");
                }

                map._byEntityTin[tin] = new Entry
                {
                    EntityTin = tin,
                    Country = countryCode,
                    SubGroupTypesRaw = sgType,
                    SubGroupTin = string.IsNullOrEmpty(sgTin) ? null : sgTin,
                };
            }
            return map;
        }

        public bool TryGet(string entityTin, out Entry entry) =>
            _byEntityTin.TryGetValue(entityTin, out entry);

        public IEnumerable<Entry> Entries => _byEntityTin.Values;
    }
}
