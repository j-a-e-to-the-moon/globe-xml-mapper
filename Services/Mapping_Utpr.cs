using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// UTPR 배분 시트 → GlobeBody.UtprAttribution 매핑.
    /// 컬럼 헤더("1.소득산입보완규칙") 다음 행부터 데이터로 iteration. 합계 행 자동 스킵.
    /// </summary>
    public class Mapping_Utpr : MappingBase
    {
        private const string COLUMN_HEADER_ANCHOR = "1.소득산입보완규칙"; // B열 컬럼헤더 (점 뒤 공백 없음)

        public Mapping_Utpr()
            : base(null) { }

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            var headerRow = FindAnchorRow(ws, COLUMN_HEADER_ANCHOR);
            if (headerRow < 0) return;
            var dataStartRow = headerRow + 1;
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? dataStartRow;

            var utpr = new Globe.GlobeBodyTypeUtprAttribution();

            // RecJurCode: 신고구성기업 소재지국
            var filingCountry = globe.GlobeBody.FilingInfo?.FilingCe?.ResCountryCode;
            if (filingCountry.HasValue)
                utpr.RecJurCode.Add(filingCountry.Value);

            bool hasData = false;

            for (int row = dataStartRow; row <= lastRow; row++)
            {
                var resCodeRaw = ws.Cell(row, 2).GetString()?.Trim(); // B
                if (string.IsNullOrEmpty(resCodeRaw))
                    continue;

                // 합계 행 등 국가코드 파싱 불가 → 스킵
                if (!TryParseEnum<Globe.CountryCodeType>(resCodeRaw, out var country))
                    continue;

                // [R] xs:integer 필드는 빈 셀이면 "0", [O] 필드는 NullIfEmpty
                var attr = new Globe.UtprAttributionTypeAttribution
                {
                    ResCountryCode = country,
                    UtprTopUpTaxCarryForward = ws.Cell(row, 3).GetString()?.Trim() ?? "0", // C [R]
                    Employees = NullIfEmpty(ws.Cell(row, 5).GetString()?.Trim()), // E [O]
                    TangibleAssetValue = NullIfEmpty(ws.Cell(row, 7).GetString()?.Trim()), // G [O]
                    UtprTopUpTaxAttributed = ws.Cell(row, 12).GetString()?.Trim() ?? "0", // L [R]
                    AddCashTaxExpense = ws.Cell(row, 14).GetString()?.Trim() ?? "0", // N [R]
                    UtprTopUpTaxCarriedForward = ws.Cell(row, 17).GetString()?.Trim() ?? "0", // Q [R]
                };

                // J(10): 배분율 (0~1 decimal; 퍼센트 입력 시 /100)
                var pct = ParsePercentage(ws.Cell(row, 10).GetString()?.Trim());
                if (pct.HasValue)
                    attr.UtprPercentage = pct.Value;

                utpr.Attribution.Add(attr);
                hasData = true;
            }

            if (hasData)
                globe.GlobeBody.UtprAttribution.Add(utpr);
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
