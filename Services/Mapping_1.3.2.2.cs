using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_3_2_2 : MappingBase
    {
        // 동적 탐지: 첫 블록은 "1.3.2.2 제외기업" 헤더로 찾음.
        // 후속 블록은 SET_SIZE 간격(헤더 1 + 필드 3 + 빈 2 = 6행)으로 반복.
        private const string BLOCK_ANCHOR = "1.3.2.2 제외기업";
        private const int BLOCK_SIZE = 4; // 헤더 + 3 필드 = 4행
        private const int BLOCK_GAP = 2; // 블록 간 구분 2행
        private const int SET_SIZE = BLOCK_SIZE + BLOCK_GAP; // = 6

        // 블록 내 상대 오프셋 (블록 시작 행 기준, O열 = 15)
        // +0 = 헤더행("1.3.2.2 제외기업"), 데이터 없음
        // +1 = 1. 변동 여부
        // +2 = 2. 제외기업 상호
        // +3 = 3. 제외기업 유형
        private static readonly (int Offset, string Target)[] FieldMap =
        {
            (1, "Change"),
            (2, "Name"),
            (3, "Type"),
        };

        public Mapping_1_3_2_2()
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

            var firstBlockStart = FindAnchorRow(ws, BLOCK_ANCHOR);
            if (firstBlockStart < 0) return;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int n = 0; ; n++)
            {
                var blockStart = firstBlockStart + n * SET_SIZE;
                if (blockStart > lastRow)
                    break;

                var entity = new Globe.CorporateStructureTypeExcludedEntity();
                bool hasData = false;

                foreach (var (offset, target) in FieldMap)
                {
                    var row = blockStart + offset;
                    var val = ws.Cell(row, 15).GetString()?.Trim();
                    if (string.IsNullOrEmpty(val))
                        continue;

                    // Change(변동여부)만으로는 실제 데이터로 간주하지 않음 — 템플릿 기본값("부") 회피.
                    // Name 또는 Type 중 하나라도 있어야 실제 제외기업 엔트리.
                    if (target != "Change")
                        hasData = true;
                    var entry = new MappingEntry { Cell = $"O{row}", Label = target };

                    switch (target)
                    {
                        case "Change":
                            entity.Change = ParseBool(val);
                            break;
                        case "Name":
                            var (eName, eKName) = ParseNameKName(val);
                            entity.Name = eName;
                            if (eKName != null)
                                entity.KName = eKName;
                            break;
                        case "Type":
                            SetEnum<Globe.ExcludedEntityEnumType>(
                                val,
                                v => entity.Type = v,
                                errors,
                                fileName,
                                entry
                            );
                            break;
                    }
                }

                if (hasData)
                    globe.GlobeBody.GeneralSection.CorporateStructure.ExcludedEntity.Add(entity);
            }
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
