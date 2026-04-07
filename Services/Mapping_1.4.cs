using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 1.4 글로벌최저한세 정보 요약 — Summary Collection에 매핑.
    /// 행 반복 방식: 4행부터 데이터, blockCount로 행 수 결정.
    /// </summary>
    public class Mapping_1_4 : MappingBase
    {
        private const int DATA_START_ROW = 4;

        public Mapping_1_4() : base("mapping_1.4.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            var rowCount = 1;
            if (ws.Workbook.TryGetWorksheet(ExcelController.MetaSheetName, out var metaWs))
                rowCount = ExcelController.ReadBlockCount(metaWs, ws.Name);

            for (int i = 0; i < rowCount; i++)
            {
                var row = DATA_START_ROW + i;
                var summary = new Globe.GlobeBodyTypeSummary();
                summary.Jurisdiction = new Globe.SummaryTypeJurisdiction();

                // B: 소재지국
                var jurCode = ws.Cell(row, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(jurCode))
                {
                    SetEnum<Globe.CountryCodeType>(jurCode, v =>
                    {
                        summary.Jurisdiction.JurisdictionName = v;
                        summary.Jurisdiction.JurisdictionNameSpecified = true;
                    }, errors, fileName, new MappingEntry { Cell = $"B{row}", Label = "소재지국" });
                }

                // I-J: 적용면제/제외 사유 (SafeHarbour — Collection)
                var safeHarbour = ws.Cell(row, 9).GetString()?.Trim();
                if (!string.IsNullOrEmpty(safeHarbour))
                    SetEnum<Globe.SafeHarbourEnumType>(safeHarbour, v => summary.SafeHarbour.Add(v),
                        errors, fileName, new MappingEntry { Cell = $"I{row}", Label = "적용면제" });

                // K-L: 실효세율 범위
                var etrRange = ws.Cell(row, 11).GetString()?.Trim();
                if (!string.IsNullOrEmpty(etrRange))
                {
                    SetEnum<Globe.EtrRangeEnumType>(etrRange, v =>
                    {
                        summary.EtrRange = v;
                        summary.EtrRangeSpecified = true;
                    }, errors, fileName, new MappingEntry { Cell = $"K{row}", Label = "실효세율범위" });
                }

                // O-P: 추가세액(QDMTT) 범위
                var qdmtt = ws.Cell(row, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(qdmtt))
                {
                    SetEnum<Globe.QdmtTuTEnumType>(qdmtt, v =>
                    {
                        summary.QdmtTut = v;
                        summary.QdmtTutSpecified = true;
                    }, errors, fileName, new MappingEntry { Cell = $"O{row}", Label = "QDMTT범위" });
                }

                // Q-R: 추가세액(GloBE) 범위
                var globeTut = ws.Cell(row, 17).GetString()?.Trim();
                if (!string.IsNullOrEmpty(globeTut))
                {
                    SetEnum<Globe.GlobeTuTEnumType>(globeTut, v =>
                    {
                        summary.GLoBeTut = v;
                        summary.GLoBeTutSpecified = true;
                    }, errors, fileName, new MappingEntry { Cell = $"Q{row}", Label = "GloBE범위" });
                }

                // 값이 하나라도 있으면 추가
                if (!string.IsNullOrEmpty(jurCode) || !string.IsNullOrEmpty(safeHarbour)
                    || !string.IsNullOrEmpty(etrRange))
                    globe.GlobeBody.Summary.Add(summary);
            }
        }
    }
}
