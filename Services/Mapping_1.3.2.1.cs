using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_2_1 : MappingBase
    {
        // 헤더 검색 키워드 (B열에서 이 문자열을 포함하는 행 = 블록 시작)
        private const string BLOCK_HEADER = "1.3.2.1";

        // 소유지분 통합 입력 셀 (블록 내 상대 행 오프셋)
        // 블록 시작(3행) + 8 = 11행 → O11 (병합 셀 O11:R14의 앵커)
        // 포맷: "유형,TIN,TIN유형,발급국가,지분" × 주주 수 (주주 구분: ';')
        private const int OWNERSHIP_ROW_OFFSET = 8;

        // 블록 내 상대 오프셋 (헤더 행 기준)
        // +1=변동, +2=소재지국, +3=규칙, +4=상호, +5=TIN,
        // +6=접수TIN, +7=기업유형, +8=소유지분(통합 O11:R14), +12=모기업유형, +13=QIIR TIN,
        // +14=부분소유TIN, +15=QUTPR초기, +16=소유합계, +17=UPE소유
        private static readonly (int Offset, string Target)[] FieldMap =
        {
            (1, "Ce.ChangeFlag"),
            (2, "Ce.Id.ResCountryCode"),
            (3, "Ce.Id.Rules"),
            (4, "Ce.Id.Name"),
            (5, "Ce.Id.Tin.Value"),
            (6, "Ce.Id.ReceivingTin"),
            (7, "Ce.Id.GlobeStatus"),
            (12, "Ce.Qiir.PopeIpe"),
            (13, "Ce.Qiir.Exception.Art213.Tin"), // O16: Art2.1.3 해당 모기업 TIN
            (14, "Ce.Qiir.Exception.Art215.Tin"), // O17: Art2.1.5 해당 부분소유모기업 TIN
            (15, "Ce.Qutpr.Art93"),
            (16, "Ce.Qutpr.AggOwnership"),
            (17, "Ce.Qutpr.UpeOwnership"),
        };

        public Mapping_1_3_2_1()
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

            // 헤더 행 위치를 동적으로 검색
            var blockStartRows = FindBlockStartRows(ws);

            for (int blockIdx = 0; blockIdx < blockStartRows.Count; blockIdx++)
            {
                var blockStartRow = blockStartRows[blockIdx];
                var ce = new Globe.CorporateStructureTypeCe { Id = new Globe.IdType() };
                bool hasData = false;

                // 블록 내 필드 매핑
                foreach (var (offset, target) in FieldMap)
                {
                    var row = blockStartRow + offset;
                    var cellValue = ws.Cell(row, 15).GetString()?.Trim(); // O열 = 15
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    hasData = true;
                    var isMulti = target is "Ce.Id.Rules" or "Ce.Id.GlobeStatus";
                    var values = isMulti
                        ? cellValue.Split(
                            ',',
                            System.StringSplitOptions.RemoveEmptyEntries
                                | System.StringSplitOptions.TrimEntries
                        )
                        : new[] { cellValue };

                    foreach (var val in values)
                        SetCeValue(ce, cs, target, val, errors, fileName, row);
                }

                // 소유지분 통합 셀(O12) 파싱 — 주주 여러 명을 ';'로 구분, 각 주주는 ',' 5필드
                var ownershipRow = blockStartRow + OWNERSHIP_ROW_OFFSET;
                var ownershipCell = ws.Cell(ownershipRow, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(ownershipCell))
                {
                    hasData = true;
                    MapOwnershipFromCell(ce, ownershipCell, ownershipRow, errors, fileName);
                }

                if (hasData)
                    cs.Ce.Add(ce);
            }
        }

        private void SetCeValue(
            Globe.CorporateStructureTypeCe ce,
            Globe.CorporateStructureType cs,
            string target,
            string val,
            List<string> errors,
            string fileName,
            int row
        )
        {
            var entry = new MappingEntry { Cell = $"O{row}", Label = target };

            switch (target)
            {
                case "Ce.ChangeFlag":
                    cs.UnreportChangeCorpStr = ParseBool(val);
                    cs.UnreportChangeCorpStrSpecified = true;
                    break;
                case "Ce.Id.ResCountryCode":
                    SetEnum<Globe.CountryCodeType>(
                        val,
                        v => ce.Id.ResCountryCode.Add(v),
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Id.Rules":
                    SetEnum<Globe.IdTypeRulesEnumType>(
                        val,
                        v => ce.Id.Rules.Add(v),
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Id.Name":
                    var (ceName, ceKName) = ParseNameKName(val);
                    ce.Id.Name = ceName;
                    if (ceKName != null)
                        ce.Id.KName = ceKName;
                    break;
                case "Ce.Id.Tin.Value":
                    ce.Id.Tin.Add(ParseTin(val));
                    break;
                case "Ce.Id.ReceivingTin":
                    ce.Id.Tin.Add(ParseTin(val));
                    break;
                case "Ce.Id.GlobeStatus":
                    SetEnum<Globe.IdTypeGloBeStatusEnumType>(
                        val,
                        v => ce.Id.GlobeStatus.Add(v),
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Qiir.PopeIpe":
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    SetEnum<Globe.PopeipeEnumType>(
                        val,
                        v => ce.Qiir.PopeIpe = v,
                        errors,
                        fileName,
                        entry
                    );
                    break;
                case "Ce.Qiir.Exception.Art213.Tin":
                    // Art2.1.3: 최종모기업/중간모기업에 QIIR 적용 시 해당 모기업 TIN
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    ce.Qiir.Exception ??= new Globe.CorporateStructureTypeCeQiirException();
                    ce.Qiir.Exception.ExceptionRule ??=
                        new Globe.CorporateStructureTypeCeQiirExceptionExceptionRule();
                    ce.Qiir.Exception.ExceptionRule.Art213 = true;
                    ce.Qiir.Exception.ExceptionRule.Art213Specified = true;
                    ce.Qiir.Exception.Tin = ParseTin(val);
                    break;
                case "Ce.Qiir.Exception.Art215.Tin":
                    // Art2.1.5: 부분소유모기업에 QIIR 적용 시 해당 모기업 TIN
                    ce.Qiir ??= new Globe.CorporateStructureTypeCeQiir();
                    ce.Qiir.Exception ??= new Globe.CorporateStructureTypeCeQiirException();
                    ce.Qiir.Exception.ExceptionRule ??=
                        new Globe.CorporateStructureTypeCeQiirExceptionExceptionRule();
                    ce.Qiir.Exception.ExceptionRule.Art215 = true;
                    ce.Qiir.Exception.ExceptionRule.Art215Specified = true;
                    ce.Qiir.Exception.Tin = ParseTin(val);
                    break;
                case "Ce.Qutpr.Art93":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    ce.Qutpr.Art93 = ParseBool(val);
                    break;
                case "Ce.Qutpr.AggOwnership":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    var agg = ParsePercentage(val);
                    if (agg.HasValue)
                    {
                        ce.Qutpr.AggOwnership = agg.Value;
                        ce.Qutpr.AggOwnershipSpecified = true;
                    }
                    break;
                case "Ce.Qutpr.UpeOwnership":
                    ce.Qutpr ??= new Globe.CorporateStructureTypeCeQutpr();
                    ce.Qutpr.UpeOwnership = ParseBool(val);
                    ce.Qutpr.UpeOwnershipSpecified = true;
                    break;
            }
        }

        /// <summary>
        /// 소유지분 통합 셀 파싱.
        /// 포맷: "유형,TIN,TIN유형,발급국가,지분" × N (주주 구분 ';')
        /// 예: "GIR801, 1234567890, GIR3001, KR, 1; GIR802, 987, GIR3001, KR, 0.5"
        /// </summary>
        private void MapOwnershipFromCell(
            Globe.CorporateStructureTypeCe ce,
            string cellValue,
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

                var ownership = new Globe.CorporateStructureTypeCeOwnership();
                var entry = new MappingEntry { Cell = $"O{row}", Label = $"소유지분[{i + 1}]" };

                // [0] 유형 (GIR801~806)
                if (parts.Length >= 1 && !string.IsNullOrEmpty(parts[0]))
                    SetEnum<Globe.OwnershipTypeEnumType>(
                        parts[0],
                        v => ownership.OwnershipType = v,
                        errors,
                        fileName,
                        entry
                    );

                // [1] TIN 값 + [2] TIN유형 (GIR300x) + [3] 발급국가 (ISO2)
                var tinValue = parts.Length >= 2 ? parts[1] : null;
                if (!string.IsNullOrEmpty(tinValue))
                {
                    var tin = new Globe.TinType { Value = tinValue };
                    if (
                        parts.Length >= 3
                        && !string.IsNullOrEmpty(parts[2])
                        && TryParseEnum<Globe.TinEnumType>(parts[2], out var tinType)
                    )
                    {
                        tin.TypeOfTin = tinType;
                        tin.TypeOfTinSpecified = true;
                    }
                    if (
                        parts.Length >= 4
                        && !string.IsNullOrEmpty(parts[3])
                        && TryParseEnum<Globe.CountryCodeType>(parts[3], out var country)
                    )
                    {
                        tin.IssuedBy = country;
                        tin.IssuedBySpecified = true;
                    }
                    ownership.Tin = tin;
                }
                else
                {
                    ownership.Tin = NoTin();
                }

                // [4] 지분 (0~1 또는 % 단위)
                if (parts.Length >= 5 && !string.IsNullOrEmpty(parts[4]))
                {
                    var pct = ParsePercentage(parts[4]);
                    if (pct.HasValue)
                        ownership.OwnershipPercentage = pct.Value;
                }

                ce.Ownership.Add(ownership);
            }
        }

        /// <summary>
        /// B열에서 BLOCK_HEADER 문자열을 포함하는 행을 모두 찾아 반환.
        /// 각 행이 CE 블록의 시작점(헤더 행).
        /// </summary>
        private static List<int> FindBlockStartRows(IXLWorksheet ws)
        {
            var result = new List<int>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 200;
            for (int r = 1; r <= lastRow; r++)
            {
                var val = ws.Cell(r, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(val) && val.Contains(BLOCK_HEADER))
                    result.Add(r);
            }
            return result;
        }
    }
}
