using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 1.3.3 기업구조 변동 — AdditionalDataPoint Collection에 매핑.
    /// 행 반복 방식: 7행부터 데이터, blockCount로 행 수 결정.
    /// </summary>
    public class Mapping_1_3_3 : MappingBase
    {
        private const int DATA_START_ROW = 7;

        public Mapping_1_3_3() : base("mapping_1.3.3.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();

            var rowCount = 1;
            if (ws.Workbook.TryGetWorksheet(ExcelController.MetaSheetName, out var metaWs))
                rowCount = ExcelController.ReadBlockCount(metaWs, ws.Name);

            for (int i = 0; i < rowCount; i++)
            {
                var row = DATA_START_ROW + i;
                var dp = new Globe.AdditionalDataPointType();

                // B-C: 상호 → Description
                var name = ws.Cell(row, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(name))
                    dp.Description = name;

                // D-E: 납세자번호 → Text (TIN을 텍스트로)
                var tin = ws.Cell(row, 4).GetString()?.Trim();

                // F-G: 변동효력발생일
                var dateVal = ws.Cell(row, 6).GetString()?.Trim();

                // H-I: 변동 전 기업유형
                var pretype = ws.Cell(row, 8).GetString()?.Trim();

                // J: 변동 후 기업유형
                var posttype = ws.Cell(row, 10).GetString()?.Trim();

                // K-L: 소유지분 보유 기업
                var owner = ws.Cell(row, 11).GetString()?.Trim();

                // M-O: 변동 전 소유지분(%)
                var prePct = ws.Cell(row, 13).GetString()?.Trim();

                // P-R: 변동 후 소유지분(%)
                var postPct = ws.Cell(row, 16).GetString()?.Trim();

                // 텍스트로 통합 저장
                var parts = new List<string>();
                if (!string.IsNullOrEmpty(tin)) parts.Add($"TIN:{tin}");
                if (!string.IsNullOrEmpty(dateVal)) parts.Add($"변동일:{dateVal}");
                if (!string.IsNullOrEmpty(pretype)) parts.Add($"변동전:{pretype}");
                if (!string.IsNullOrEmpty(posttype)) parts.Add($"변동후:{posttype}");
                if (!string.IsNullOrEmpty(owner)) parts.Add($"소유기업:{owner}");
                if (!string.IsNullOrEmpty(prePct)) parts.Add($"변동전%:{prePct}");
                if (!string.IsNullOrEmpty(postPct)) parts.Add($"변동후%:{postPct}");

                if (parts.Count > 0)
                    dp.Text = string.Join("; ", parts);

                // 값이 하나라도 있으면 추가
                if (!string.IsNullOrEmpty(dp.Description) || !string.IsNullOrEmpty(dp.Text))
                    globe.GlobeBody.GeneralSection.AdditionalDataPoint.Add(dp);
            }
        }
    }
}
