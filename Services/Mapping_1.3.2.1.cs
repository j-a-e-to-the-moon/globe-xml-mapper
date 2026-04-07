using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_2_1 : MappingBase
    {
        // CE 블록 정의 (시트2 기준)
        private const int BLOCK_START = 4;
        private const int BLOCK_END = 21;
        private const int BLOCK_GAP = 2;
        private const int BLOCK_SIZE = BLOCK_END - BLOCK_START + 1; // 18행

        // 블록 내 상대 오프셋 (행 번호 - BLOCK_START)
        // O5=변동(+1), O6=소재지국(+2), O7=규칙(+3), O8=상호(+4), O9=TIN(+5),
        // O10=접수TIN(+6), O11=기업유형(+7), O14=별첨참조(+10),
        // O16=모기업유형(+12), O17=QIIR TIN(+13), O18=부분소유TIN(+14),
        // O19=QUTPR초기(+15), O20=소유합계(+16), O21=UPE소유(+17)
        private static readonly (int Offset, string Target)[] FieldMap =
        {
            (1,  "Ce.ChangeFlag"),
            (2,  "Ce.Id.ResCountryCode"),
            (3,  "Ce.Id.Rules"),
            (4,  "Ce.Id.Name"),
            (5,  "Ce.Id.Tin.Value"),
            (6,  "Ce.Id.ReceivingTin"),
            (7,  "Ce.Id.GlobeStatus"),
            (12, "Ce.Qiir.PopeIpe"),
            (13, "Ce.Qiir.Exception.Tin.Value"),
            (14, "Ce.Qiir.MopeIpe.Tin.Value"),
            (15, "Ce.Qutpr.Art93"),
            (16, "Ce.Qutpr.AggOwnership"),
            (17, "Ce.Qutpr.UpeOwnership"),
        };

        // 별첨 시트 이름
        private const string ATTACH_SHEET = "부표2 (2) 별첨";

        public Mapping_1_3_2_1() : base("mapping_1.3.2.1.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??= new Globe.CorporateStructureType();
            var cs = globe.GlobeBody.GeneralSection.CorporateStructure;

            // _META에서 CE 블록 수 읽기
            var blockCount = 1;
            if (ws.Workbook.TryGetWorksheet(ExcelController.MetaSheetName, out var metaWs))
                blockCount = ExcelController.ReadBlockCount(metaWs, ws.Name);

            // CE 블록 순회
            for (int blockIdx = 0; blockIdx < blockCount; blockIdx++)
            {
                var blockStartRow = BLOCK_START + blockIdx * (BLOCK_SIZE + BLOCK_GAP);
                var ce = new Globe.CorporateStructureTypeCe { Id = new Globe.IdType() };
                cs.Ce.Add(ce);

                // 블록 내 필드 매핑
                foreach (var (offset, target) in FieldMap)
                {
                    var row = blockStartRow + offset;
                    var cellValue = ws.Cell(row, 15).GetString()?.Trim(); // O열 = 15
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    // multi 처리 (쉼표 구분)
                    var isMulti = target is "Ce.Id.Rules" or "Ce.Id.GlobeStatus";
                    var values = isMulti
                        ? cellValue.Split(',', System.StringSplitOptions.RemoveEmptyEntries | System.StringSplitOptions.TrimEntries)
                        : new[] { cellValue };

                    foreach (var val in values)
                        SetCeValue(ce, cs, target, val, errors, fileName, row);
                }

                // 별첨에서 소유지분 읽기
                var attachNum = blockIdx + 1;
                MapOwnershipFromAttach(ws.Workbook, ce, attachNum, errors, fileName);
            }
        }

        private void SetCeValue(Globe.CorporateStructureTypeCe ce, Globe.CorporateStructureType cs,
            string target, string val, List<string> errors, string fileName, int row)
        {
            var entry = new MappingEntry { Cell = $"O{row}", Label = target };

            switch (target)
            {
                case "Ce.ChangeFlag":
                    cs.UnreportChangeCorpStr = ParseBool(val);
                    cs.UnreportChangeCorpStrSpecified = true; break;
                case "Ce.Id.ResCountryCode":
                    SetEnum<Globe.CountryCodeType>(val, v => ce.Id.ResCountryCode.Add(v), errors, fileName, entry); break;
                case "Ce.Id.Rules":
                    SetEnum<Globe.IdTypeRulesEnumType>(val, v => ce.Id.Rules.Add(v), errors, fileName, entry); break;
                case "Ce.Id.Name": ce.Id.Name = val; break;
                case "Ce.Id.Tin.Value":
                    ce.Id.Tin.Add(new Globe.TinType { Value = val }); break;
                case "Ce.Id.ReceivingTin":
                    ce.Id.Tin.Add(new Globe.TinType { Value = val, IssuedBy = Globe.CountryCodeType.Kr, IssuedBySpecified = true }); break;
                case "Ce.Id.GlobeStatus":
                    SetEnum<Globe.IdTypeGloBeStatusEnumType>(val, v => ce.Id.GlobeStatus.Add(v), errors, fileName, entry); break;
                case "Ce.Qiir.PopeIpe":
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    SetEnum<Globe.PopeipeEnumType>(val, v => ce.Qiir.PopeIpe = v, errors, fileName, entry); break;
                case "Ce.Qiir.Exception.Tin.Value":
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    ce.Qiir.Exception ??= new Globe.CorporateStructureTypeCeQiirException();
                    ce.Qiir.Exception.Tin = new Globe.TinType { Value = val }; break;
                case "Ce.Qiir.MopeIpe.Tin.Value": break;
                case "Ce.Qutpr.Art93":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    ce.Qutpr.Art93 = ParseBool(val); break;
                case "Ce.Qutpr.AggOwnership":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    if (decimal.TryParse(val, out var agg))
                    { ce.Qutpr.AggOwnership = agg / 100m; ce.Qutpr.AggOwnershipSpecified = true; }
                    break;
                case "Ce.Qutpr.UpeOwnership":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    ce.Qutpr.UpeOwnership = ParseBool(val);
                    ce.Qutpr.UpeOwnershipSpecified = true; break;
            }
        }

        /// <summary>
        /// 별첨 시트에서 별첨N의 주주 데이터를 읽어 Ownership에 추가.
        /// </summary>
        private void MapOwnershipFromAttach(IXLWorkbook workbook,
            Globe.CorporateStructureTypeCe ce, int attachNum,
            List<string> errors, string fileName)
        {
            if (!workbook.TryGetWorksheet(ATTACH_SHEET, out var attachWs))
                return;

            // "별첨N" 제목 행 찾기
            var startRow = FindAttachStart(attachWs, attachNum);
            if (startRow < 0) return;

            // 제목(1) + 빈행(1) + 헤더(1) = 3행 뒤부터 데이터
            var dataRow = startRow + 3;
            while (true)
            {
                var typeVal = attachWs.Cell(dataRow, 2).GetString()?.Trim();
                var tinVal = attachWs.Cell(dataRow, 3).GetString()?.Trim();
                var pctVal = attachWs.Cell(dataRow, 4).GetString()?.Trim();

                if (string.IsNullOrEmpty(typeVal) && string.IsNullOrEmpty(tinVal) && string.IsNullOrEmpty(pctVal))
                    break;

                // 다음 별첨 제목이면 종료
                if (typeVal != null && typeVal.StartsWith("별첨")) break;

                var ownership = new Globe.CorporateStructureTypeCeOwnership();

                if (!string.IsNullOrEmpty(typeVal))
                    SetEnum<Globe.OwnershipTypeEnumType>(typeVal, v => ownership.OwnershipType = v,
                        errors, fileName, new MappingEntry { Cell = $"별첨!B{dataRow}", Label = $"별첨{attachNum} 유형" });

                if (!string.IsNullOrEmpty(tinVal))
                    ownership.Tin = new Globe.TinType { Value = tinVal };

                if (!string.IsNullOrEmpty(pctVal) && decimal.TryParse(pctVal, out var pct))
                    ownership.OwnershipPercentage = pct / 100m;

                ce.Ownership.Add(ownership);
                dataRow++;
            }
        }

        private static int FindAttachStart(IXLWorksheet ws, int attachNum)
        {
            var target = $"별첨{attachNum}";
            for (int r = 1; r <= 500; r++)
            {
                var val = ws.Cell(r, 2).GetString()?.Trim();
                if (val == target) return r;
            }
            return -1;
        }
    }
}
