using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace GlobeMapper.Services
{
    /// <summary>
    /// Excel COM late-binding 래퍼.
    /// 시트 복제(CE/제외기업) + 시트 내 행 블록 반복(UPE) 지원.
    /// </summary>
    public class ExcelController : IDisposable
    {
        private dynamic _app;
        private dynamic _workbook;
        private bool _disposed;

        public const string MetaSheetName = "_META";

        public event Action WorkbookClosed;

        public bool IsOpen
        {
            get
            {
                try { return _workbook != null && _app?.Visible == true; }
                catch { return false; }
            }
        }

        public string GetActiveSheetName()
        {
            try { return (string)_app?.ActiveSheet?.Name; }
            catch { return null; }
        }

        public void ActivateSheet(object sheetNameOrIndex)
        {
            try { _workbook.Sheets[sheetNameOrIndex].Activate(); }
            catch { }
        }

        public void Open(string path)
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Excel이 설치되어 있지 않습니다.");

            _app = Activator.CreateInstance(excelType);
            _app.Visible = true;
            _workbook = _app.Workbooks.Open(path);

            // 첫 번째 시트로 이동 (메타 시트 생성 전 위치 기억)
            var firstSheet = _workbook.Sheets[1];
            EnsureMetaSheet();
            // 메타 시트 생성 후 원래 시트로 복귀
            ((dynamic)firstSheet).Activate();
        }

        public void CreateNew(string templatePath, string savePath)
        {
            System.IO.File.Copy(templatePath, savePath, true);
            Open(savePath);
        }

        public void Save() => _workbook?.Save();

        public string GetFilePathForMapping()
        {
            Save();
            return (string)_workbook.FullName;
        }

        public void CloseWithSavePrompt()
        {
            if (_workbook == null) return;
            try
            {
                bool saved = (bool)_workbook.Saved;
                if (!saved)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "변경사항이 있습니다. 저장하시겠습니까?",
                        "저장 확인",
                        System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                        System.Windows.Forms.MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Cancel) return;
                    _workbook.Close(SaveChanges: result == System.Windows.Forms.DialogResult.Yes);
                }
                else
                {
                    _workbook.Close(SaveChanges: false);
                }
            }
            catch { }
            finally { QuitApp(); }
        }

        #region 시트 내 행 블록 반복 (1.3.1 UPE)

        /// <summary>
        /// 시트 내 행 블록을 복제하여 아래에 추가.
        /// sourceStartRow~sourceEndRow를 복사하여 현재 마지막 블록 + gap행 뒤에 삽입.
        /// </summary>
        public void AddRowBlock(string sheetName, int sourceStartRow, int sourceEndRow, int gap)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var blockSize = sourceEndRow - sourceStartRow + 1;
            var count = GetMetaInt(sheetName, "blockCount", 1);
            var insertRow = sourceEndRow + 1 + (count - 1) * (blockSize + gap) + gap;

            // 빈 행 삽입
            dynamic insertRange = ws.Rows[$"{insertRow}:{insertRow + blockSize - 1}"];
            insertRange.Insert();

            // 원본 블록 복사
            dynamic sourceRange = ws.Range[
                ws.Cells[sourceStartRow, 1],
                ws.Cells[sourceEndRow, 18]  // R열 = 18
            ];
            dynamic destRange = ws.Range[
                ws.Cells[insertRow, 1],
                ws.Cells[insertRow + blockSize - 1, 18]
            ];
            sourceRange.Copy(destRange);

            // 행 높이 복사
            for (int i = 0; i < blockSize; i++)
            {
                ws.Rows[insertRow + i].RowHeight = (double)ws.Rows[sourceStartRow + i].RowHeight;
            }

            // 데이터 셀 초기화 (값만 지우기, 서식 유지)
            ClearDataCells(ws, insertRow, insertRow + blockSize - 1);

            SetMetaInt(sheetName, "blockCount", count + 1);
        }

        /// <summary>
        /// 마지막 행 블록 삭제.
        /// </summary>
        public bool RemoveRowBlock(string sheetName, int sourceStartRow, int sourceEndRow, int gap)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count <= 1) return false;

            dynamic ws = _workbook.Sheets[sheetName];
            var blockSize = sourceEndRow - sourceStartRow + 1;
            var lastBlockStart = sourceEndRow + 1 + (count - 2) * (blockSize + gap) + gap;
            var lastBlockEnd = lastBlockStart + blockSize - 1;

            // gap행 포함 삭제
            _app.DisplayAlerts = false;
            try
            {
                dynamic deleteRange = ws.Rows[$"{lastBlockStart - gap}:{lastBlockEnd}"];
                deleteRange.Delete();
            }
            finally
            {
                _app.DisplayAlerts = true;
            }

            SetMetaInt(sheetName, "blockCount", count - 1);
            return true;
        }

        /// <summary>
        /// 시트를 원래 상태로 초기화 (추가된 블록 모두 제거 + 데이터 초기화).
        /// </summary>
        public void ResetSheet(string sheetName, int sourceStartRow, int sourceEndRow, int gap)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);

            if (count > 1)
            {
                dynamic ws = _workbook.Sheets[sheetName];
                var blockSize = sourceEndRow - sourceStartRow + 1;
                var firstExtraRow = sourceEndRow + 1 + gap;
                var lastRow = sourceEndRow + (count - 1) * (blockSize + gap);

                _app.DisplayAlerts = false;
                try
                {
                    dynamic deleteRange = ws.Rows[$"{firstExtraRow}:{lastRow}"];
                    deleteRange.Delete();
                }
                finally
                {
                    _app.DisplayAlerts = true;
                }
            }

            // 원본 블록 데이터도 초기화
            dynamic sheet = _workbook.Sheets[sheetName];
            ClearDataCells(sheet, sourceStartRow, sourceEndRow);

            SetMetaInt(sheetName, "blockCount", 1);
        }

        public int GetRowBlockCount(string sheetName)
        {
            return GetMetaInt(sheetName, "blockCount", 1);
        }

        private void ClearDataCells(dynamic ws, int startRow, int endRow)
        {
            // O열~R열 (15~18)의 데이터 셀만 값 초기화 (서식 유지)
            for (int r = startRow; r <= endRow; r++)
            {
                for (int c = 15; c <= 18; c++)
                {
                    dynamic cell = ws.Cells[r, c];
                    if (cell.MergeCells)
                    {
                        // 병합 셀의 첫 번째 셀만 처리
                        dynamic mergeArea = cell.MergeArea;
                        if ((int)mergeArea.Row == r && (int)mergeArea.Column == c)
                            mergeArea.ClearContents();
                    }
                    else
                    {
                        cell.ClearContents();
                    }
                }
            }
        }

        #endregion

        #region 시트 복제 (CE, 제외기업)

        private static readonly Dictionary<string, int> SheetTemplateIndex = new()
        {
            { "1.3.2.1", 2 },
            { "1.3.2.2", 3 },
        };

        public string AddSheet(string section)
        {
            if (!SheetTemplateIndex.TryGetValue(section, out var templateIdx))
                throw new ArgumentException($"알 수 없는 섹션: {section}");

            dynamic sourceSheet = _workbook.Sheets[templateIdx];
            dynamic lastSheet = _workbook.Sheets[_workbook.Sheets.Count];
            sourceSheet.Copy(After: lastSheet);

            dynamic newSheet = _workbook.Sheets[_workbook.Sheets.Count];
            var count = GetSheetCount(section);
            var newName = $"{section} ({count + 1})";
            newSheet.Name = newName;

            AddMetaEntry(section, newName);
            return newName;
        }

        public bool RemoveSheet(string section)
        {
            var sheets = GetSectionSheets(section);
            if (sheets.Count <= 1) return false;

            var lastSheetName = sheets.Last();
            _app.DisplayAlerts = false;
            try { _workbook.Sheets[lastSheetName].Delete(); }
            finally { _app.DisplayAlerts = true; }

            RemoveMetaEntry(section, lastSheetName);
            return true;
        }

        public List<string> GetSectionSheets(string section)
        {
            dynamic meta = GetMetaSheet();
            if (meta == null) return new List<string>();

            var result = new List<string>();
            var row = 2;
            while (true)
            {
                string sec = meta.Cells[row, 1].Value?.ToString();
                if (string.IsNullOrEmpty(sec)) break;
                string name = meta.Cells[row, 2].Value?.ToString();
                if (sec == section && !string.IsNullOrEmpty(name))
                    result.Add(name);
                row++;
            }
            return result;
        }

        public int GetSheetCount(string section) => GetSectionSheets(section).Count;

        #endregion

        #region CE 블록 + 첨부 시트 연동

        private const int CE_BLOCK_START = 4;
        private const int CE_BLOCK_END = 21;
        private const int CE_BLOCK_GAP = 2;
        private const int CE_ATTACH_REF_ROW_OFFSET = 10; // 블록 내 O14 = 시작행+10

        /// <summary>
        /// CE 블록 추가: 시트2에서 행 블록 복제 + 별첨 시트에 섹션 추가.
        /// </summary>
        public void AddCeBlock(string ceSheetName, string attachSheetName)
        {
            // 1. 행 블록 복제
            AddRowBlock(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP);

            var count = GetRowBlockCount(ceSheetName);
            var blockSize = CE_BLOCK_END - CE_BLOCK_START + 1;

            // 2. 새 블록의 O14셀을 "별첨N"으로 갱신
            dynamic ws = _workbook.Sheets[ceSheetName];
            var newBlockStart = CE_BLOCK_END + 1 + (count - 2) * (blockSize + CE_BLOCK_GAP) + CE_BLOCK_GAP;
            var refRow = newBlockStart + CE_ATTACH_REF_ROW_OFFSET;
            ws.Cells[refRow, 15] = $"첨부{count}"; // O열 = 15

            // 3. 별첨 시트에 섹션 추가
            AddAttachSection(attachSheetName, count);
        }

        /// <summary>
        /// 마지막 CE 블록 삭제 + 별첨 시트에서 해당 섹션 삭제.
        /// </summary>
        public bool RemoveCeBlock(string ceSheetName, string attachSheetName)
        {
            var count = GetRowBlockCount(ceSheetName);
            if (count <= 1) return false;

            RemoveRowBlock(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP);
            RemoveAttachSection(attachSheetName, count);
            return true;
        }

        /// <summary>
        /// CE 시트 + 별첨 시트 초기화.
        /// </summary>
        public void ResetCeSheet(string ceSheetName, string attachSheetName)
        {
            var count = GetRowBlockCount(ceSheetName);

            // 별첨 시트 초기화: 별첨2 이후 모두 삭제
            if (count > 1)
            {
                for (int i = count; i >= 2; i--)
                    RemoveAttachSection(attachSheetName, i);
            }
            // 별첨1 데이터 행 초기화
            ResetAttachSection(attachSheetName, 1);

            // CE 시트 초기화
            ResetSheet(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP);
        }

        public int GetCeBlockCount(string ceSheetName) => GetRowBlockCount(ceSheetName);

        #endregion

        #region 별첨 시트 관리

        // 별첨 섹션 구조: 제목행(1) + 빈행(1) + 헤더행(1) + 데이터행(N) + 구분빈행(1)
        private const int ATTACH_HEADER_ROWS = 3; // 제목 + 빈행 + 헤더
        private const int ATTACH_SEPARATOR = 1;   // 구분 빈행
        private const int ATTACH_INITIAL_DATA_ROWS = 1; // 초기 데이터 행 수

        /// <summary>
        /// 별첨 시트에서 별첨N 섹션의 시작 행을 찾음.
        /// </summary>
        private int FindAttachSectionStart(dynamic ws, int attachNum)
        {
            var row = 1;
            var target = $"첨부{attachNum}";
            for (int r = 1; r <= 500; r++)
            {
                string val = ws.Cells[r, 2].Value?.ToString()?.Trim();
                if (val == target) return r;
            }
            return -1;
        }

        /// <summary>
        /// 별첨 시트에서 별첨N 섹션의 데이터 행 수를 반환.
        /// </summary>
        public int GetOwnerRowCount(string attachSheetName, int attachNum)
        {
            dynamic ws = _workbook.Sheets[attachSheetName];
            var start = FindAttachSectionStart(ws, attachNum);
            if (start < 0) return 0;

            var dataStart = start + ATTACH_HEADER_ROWS;
            var count = 0;
            for (int r = dataStart; r <= dataStart + 200; r++)
            {
                string b = ws.Cells[r, 2].Value?.ToString()?.Trim();

                // 다음 별첨 제목이면 종료
                if (b != null && b.StartsWith("첨부")) break;

                // 값이 있거나 테두리가 있으면 데이터 행으로 카운트
                string c = ws.Cells[r, 3].Value?.ToString()?.Trim();
                string d = ws.Cells[r, 4].Value?.ToString()?.Trim();
                bool hasValue = !string.IsNullOrEmpty(b) || !string.IsNullOrEmpty(c) || !string.IsNullOrEmpty(d);

                // 테두리 확인 (B열 기준)
                bool hasBorder = false;
                try
                {
                    dynamic borders = ws.Cells[r, 2].Borders;
                    // xlEdgeBottom = 9
                    hasBorder = borders[9].LineStyle != -4142; // -4142 = xlNone
                }
                catch { }

                if (!hasValue && !hasBorder) break;
                count++;
            }
            return count;
        }

        /// <summary>
        /// 별첨 시트에서 별첨N에 주주 행 1개 추가.
        /// </summary>
        public void AddOwnerRow(string attachSheetName, int attachNum)
        {
            dynamic ws = _workbook.Sheets[attachSheetName];
            var start = FindAttachSectionStart(ws, attachNum);
            if (start < 0) return;

            var dataStart = start + ATTACH_HEADER_ROWS;
            var rowCount = GetOwnerRowCount(attachSheetName, attachNum);
            var insertRow = dataStart + rowCount;

            // 첫 데이터 행(테두리 템플릿)을 복사하여 삽입
            dynamic templateRow = ws.Rows[dataStart];
            templateRow.Copy();
            ws.Rows[insertRow].Insert();
            // 삽입된 행에 붙여넣기 (서식만)
            dynamic destRow = ws.Rows[insertRow];
            destRow.PasteSpecial(-4122); // xlPasteFormats = -4122
            // 값 초기화
            ws.Cells[insertRow, 2].ClearContents();
            ws.Cells[insertRow, 3].ClearContents();
            ws.Cells[insertRow, 4].ClearContents();
            _app.CutCopyMode = false;
        }

        /// <summary>
        /// 별첨 시트에서 별첨N의 마지막 주주 행 삭제.
        /// </summary>
        public bool RemoveOwnerRow(string attachSheetName, int attachNum)
        {
            var rowCount = GetOwnerRowCount(attachSheetName, attachNum);
            if (rowCount <= 0) return false;

            dynamic ws = _workbook.Sheets[attachSheetName];
            var start = FindAttachSectionStart(ws, attachNum);
            var lastDataRow = start + ATTACH_HEADER_ROWS + rowCount - 1;

            _app.DisplayAlerts = false;
            try { ws.Rows[lastDataRow].Delete(); }
            finally { _app.DisplayAlerts = true; }
            return true;
        }

        /// <summary>
        /// 별첨 시트에 새 별첨N 섹션 추가.
        /// </summary>
        private void AddAttachSection(string attachSheetName, int attachNum)
        {
            dynamic ws = _workbook.Sheets[attachSheetName];

            // 별첨1의 헤더행+데이터행 위치 (서식 복사용)
            var attach1Start = FindAttachSectionStart(ws, 1);
            int headerRow = attach1Start >= 0 ? attach1Start + 2 : -1; // 헤더행 (유형/납세자번호/소유지분)
            int templateDataRow = attach1Start >= 0 ? attach1Start + ATTACH_HEADER_ROWS : -1; // 첫 데이터행

            // 마지막 사용 행 찾기
            int lastRow = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count;

            var startRow = lastRow + 1; // 1행 간격
            ws.Cells[startRow, 2] = $"첨부{attachNum}";

            // 헤더행: 별첨1의 헤더행 서식 복사
            if (headerRow > 0)
            {
                dynamic srcHeader = ws.Rows[headerRow];
                srcHeader.Copy();
                ws.Rows[startRow + 2].PasteSpecial(-4104); // xlPasteAll
            }
            else
            {
                ws.Cells[startRow + 2, 2] = "유형";
                ws.Cells[startRow + 2, 3] = "납세자번호";
                ws.Cells[startRow + 2, 4] = "소유지분(%)";
            }

            // 데이터행 1개: 별첨1의 첫 데이터행 서식 복사 (값은 비움)
            if (templateDataRow > 0)
            {
                dynamic srcData = ws.Rows[templateDataRow];
                srcData.Copy();
                ws.Rows[startRow + 3].PasteSpecial(-4122); // xlPasteFormats
            }

            _app.CutCopyMode = false;
        }

        /// <summary>
        /// 별첨 시트에서 마지막 별첨 섹션 삭제.
        /// </summary>
        private void RemoveAttachSection(string attachSheetName, int attachNum)
        {
            dynamic ws = _workbook.Sheets[attachSheetName];
            var start = FindAttachSectionStart(ws, attachNum);
            if (start < 0) return;

            // 해당 섹션 끝 찾기: 다음 "별첨" 또는 사용범위 끝
            int end = start;
            for (int r = start + 1; r <= start + 200; r++)
            {
                string val = ws.Cells[r, 2].Value?.ToString()?.Trim();
                if (val != null && val.StartsWith("첨부") && val != $"첨부{attachNum}")
                {
                    end = r - 1;
                    break;
                }
                end = r;
            }

            _app.DisplayAlerts = false;
            try { ws.Rows[$"{start}:{end}"].Delete(); }
            finally { _app.DisplayAlerts = true; }
        }

        /// <summary>
        /// 별첨1 데이터만 초기화 (구조 유지).
        /// </summary>
        private void ResetAttachSection(string attachSheetName, int attachNum)
        {
            dynamic ws = _workbook.Sheets[attachSheetName];
            var start = FindAttachSectionStart(ws, attachNum);
            if (start < 0) return;

            var dataStart = start + ATTACH_HEADER_ROWS;
            var rowCount = GetOwnerRowCount(attachSheetName, attachNum);
            if (rowCount > 0)
            {
                _app.DisplayAlerts = false;
                try { ws.Rows[$"{dataStart}:{dataStart + rowCount - 1}"].Delete(); }
                finally { _app.DisplayAlerts = true; }
            }
        }

        #endregion

        #region 시트3 대형 블록 (3~228, 페이지번호 행 제외)

        // 페이지번호 행 (복사에서 제외)
        private static readonly int[] S3_PAGE_ROWS = { 2, 31, 64, 92, 119, 141, 162, 184, 205 };
        private const int S3_BLOCK_START = 3;
        private const int S3_BLOCK_END = 228;
        private const int S3_PAGE_GAP = 2; // 페이지 간 간격

        public void AddSheet3Page(string sheetName)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var count = GetMetaInt(sheetName, "blockCount", 1);
            var blockSize = S3_BLOCK_END - S3_BLOCK_START + 1; // 226행

            // 삽입 위치
            var insertRow = S3_BLOCK_END + 1 + (count - 1) * (blockSize + S3_PAGE_GAP) + S3_PAGE_GAP;

            // 빈 행 삽입
            dynamic insertRange = ws.Rows[$"{insertRow}:{insertRow + blockSize - 1}"];
            insertRange.Insert();

            // 원본 전체 복사
            dynamic srcRange = ws.Range[ws.Cells[S3_BLOCK_START, 1], ws.Cells[S3_BLOCK_END, 18]];
            dynamic dstRange = ws.Range[ws.Cells[insertRow, 1], ws.Cells[insertRow + blockSize - 1, 18]];
            srcRange.Copy(dstRange);

            // 행 높이 복사
            for (int i = 0; i < blockSize; i++)
                ws.Rows[insertRow + i].RowHeight = (double)ws.Rows[S3_BLOCK_START + i].RowHeight;

            // 페이지번호 행 삭제 (새 블록 내에서)
            foreach (var pageRow in S3_PAGE_ROWS)
            {
                var offset = pageRow - S3_BLOCK_START;
                if (offset >= 0 && offset < blockSize)
                {
                    var targetRow = insertRow + offset;
                    // 페이지번호 셀(R열) 값만 삭제
                    ws.Cells[targetRow, 18].ClearContents();
                }
            }

            // 데이터 셀 초기화
            ClearDataCells(ws, insertRow, insertRow + blockSize - 1);

            SetMetaInt(sheetName, "blockCount", count + 1);
        }

        public bool RemoveSheet3Page(string sheetName)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count <= 1) return false;

            dynamic ws = _workbook.Sheets[sheetName];
            var blockSize = S3_BLOCK_END - S3_BLOCK_START + 1;
            var lastStart = S3_BLOCK_END + 1 + (count - 2) * (blockSize + S3_PAGE_GAP) + S3_PAGE_GAP;
            var lastEnd = lastStart + blockSize - 1;

            _app.DisplayAlerts = false;
            try { ws.Rows[$"{lastStart - S3_PAGE_GAP}:{lastEnd}"].Delete(); }
            finally { _app.DisplayAlerts = true; }

            SetMetaInt(sheetName, "blockCount", count - 1);
            return true;
        }

        public void ResetSheet3(string sheetName)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count > 1)
            {
                dynamic ws = _workbook.Sheets[sheetName];
                var blockSize = S3_BLOCK_END - S3_BLOCK_START + 1;
                var firstExtra = S3_BLOCK_END + 1 + S3_PAGE_GAP;
                var lastEnd = S3_BLOCK_END + (count - 1) * (blockSize + S3_PAGE_GAP) + blockSize;

                _app.DisplayAlerts = false;
                try { ws.Rows[$"{firstExtra}:{lastEnd}"].Delete(); }
                finally { _app.DisplayAlerts = true; }
            }

            dynamic sheet = _workbook.Sheets[sheetName];
            ClearDataCells(sheet, S3_BLOCK_START, S3_BLOCK_END);
            SetMetaInt(sheetName, "blockCount", 1);
        }

        // 시트3 내부 행 추가 (통합형피지배/결손금/제89조)
        public void AddSheet3Row(string sheetName, string subKey, int firstDataRow)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var metaKey = $"{sheetName}:{subKey}";
            var count = GetMetaInt(metaKey, "blockCount", GetDefaultRowCount(subKey));
            var insertRow = firstDataRow + count;

            dynamic templateRow = ws.Rows[firstDataRow];
            templateRow.Copy();
            ws.Rows[insertRow].Insert();
            dynamic destRow = ws.Rows[insertRow];
            destRow.PasteSpecial(-4122); // xlPasteFormats
            for (int c = 2; c <= 18; c++)
            {
                try { ws.Cells[insertRow, c].ClearContents(); } catch { }
            }
            _app.CutCopyMode = false;

            SetMetaInt(metaKey, "blockCount", count + 1);
        }

        public bool RemoveSheet3Row(string sheetName, string subKey, int firstDataRow)
        {
            var metaKey = $"{sheetName}:{subKey}";
            var count = GetMetaInt(metaKey, "blockCount", GetDefaultRowCount(subKey));
            if (count <= 1) return false;

            dynamic ws = _workbook.Sheets[sheetName];
            var lastRow = firstDataRow + count - 1;

            _app.DisplayAlerts = false;
            try { ws.Rows[lastRow].Delete(); }
            finally { _app.DisplayAlerts = true; }

            SetMetaInt(metaKey, "blockCount", count - 1);
            return true;
        }

        public int GetSheet3RowCount(string sheetName, string subKey)
        {
            return GetMetaInt($"{sheetName}:{subKey}", "blockCount", GetDefaultRowCount(subKey));
        }

        private static int GetDefaultRowCount(string subKey)
        {
            return subKey switch
            {
                "cfc" => 2,    // 통합형피지배 초기 2행
                "carryback" => 4, // 결손금 소급공제 초기 4행
                "art89" => 5,  // 제89조 초기 5행
                _ => 1
            };
        }

        #endregion

        #region 시트2 복합 블록 (3~23 + 26~54)

        // 시트2는 블록1(3~23) + 간격(24~25) + 블록2(26~54) = 총 52행이 한 세트
        private const int S2_BLOCK1_START = 3;
        private const int S2_BLOCK1_END = 23;
        private const int S2_GAP_ROWS = 2;  // 24~25행 (간격)
        private const int S2_BLOCK2_START = 26;
        private const int S2_BLOCK2_END = 54;
        private const int S2_TOTAL_SIZE = 52; // (23-3+1) + 2 + (54-26+1)
        private const int S2_INSERT_GAP = 2;  // 세트 간 간격

        public void AddSheet2Block(string sheetName)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var count = GetMetaInt(sheetName, "blockCount", 1);

            // 삽입 위치: 첫 세트 끝(54행) + (count-1) * (totalSize + gap) + gap
            var insertRow = S2_BLOCK2_END + 1 + (count - 1) * (S2_TOTAL_SIZE + S2_INSERT_GAP) + S2_INSERT_GAP;

            // 빈 행 삽입
            dynamic insertRange = ws.Rows[$"{insertRow}:{insertRow + S2_TOTAL_SIZE - 1}"];
            insertRange.Insert();

            // 블록1 복사 (3~23 → insertRow ~ insertRow+20)
            var block1Size = S2_BLOCK1_END - S2_BLOCK1_START + 1;
            dynamic src1 = ws.Range[ws.Cells[S2_BLOCK1_START, 1], ws.Cells[S2_BLOCK1_END, 18]];
            dynamic dst1 = ws.Range[ws.Cells[insertRow, 1], ws.Cells[insertRow + block1Size - 1, 18]];
            src1.Copy(dst1);

            // 행 높이 복사 (블록1)
            for (int i = 0; i < block1Size; i++)
                ws.Rows[insertRow + i].RowHeight = (double)ws.Rows[S2_BLOCK1_START + i].RowHeight;

            // 블록2 복사 (26~54 → insertRow+block1Size+gap ~ ...)
            var block2Start = insertRow + block1Size + S2_GAP_ROWS;
            var block2Size = S2_BLOCK2_END - S2_BLOCK2_START + 1;
            dynamic src2 = ws.Range[ws.Cells[S2_BLOCK2_START, 1], ws.Cells[S2_BLOCK2_END, 18]];
            dynamic dst2 = ws.Range[ws.Cells[block2Start, 1], ws.Cells[block2Start + block2Size - 1, 18]];
            src2.Copy(dst2);

            // 행 높이 복사 (블록2)
            for (int i = 0; i < block2Size; i++)
                ws.Rows[block2Start + i].RowHeight = (double)ws.Rows[S2_BLOCK2_START + i].RowHeight;

            // 데이터 셀 초기화
            ClearDataCells(ws, insertRow, insertRow + block1Size - 1);
            ClearDataCells(ws, block2Start, block2Start + block2Size - 1);

            SetMetaInt(sheetName, "blockCount", count + 1);
        }

        public bool RemoveSheet2Block(string sheetName)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count <= 1) return false;

            dynamic ws = _workbook.Sheets[sheetName];

            // 마지막 세트의 시작 위치
            var lastSetStart = S2_BLOCK2_END + 1 + (count - 2) * (S2_TOTAL_SIZE + S2_INSERT_GAP) + S2_INSERT_GAP;
            var lastSetEnd = lastSetStart + S2_TOTAL_SIZE - 1;

            // 간격 포함 삭제
            _app.DisplayAlerts = false;
            try
            {
                dynamic deleteRange = ws.Rows[$"{lastSetStart - S2_INSERT_GAP}:{lastSetEnd}"];
                deleteRange.Delete();
            }
            finally { _app.DisplayAlerts = true; }

            SetMetaInt(sheetName, "blockCount", count - 1);
            return true;
        }

        public void ResetSheet2(string sheetName)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count > 1)
            {
                dynamic ws = _workbook.Sheets[sheetName];
                var firstExtraStart = S2_BLOCK2_END + 1 + S2_INSERT_GAP;
                var lastEnd = S2_BLOCK2_END + (count - 1) * (S2_TOTAL_SIZE + S2_INSERT_GAP) + S2_TOTAL_SIZE;

                _app.DisplayAlerts = false;
                try { ws.Rows[$"{firstExtraStart}:{lastEnd}"].Delete(); }
                finally { _app.DisplayAlerts = true; }
            }

            dynamic sheet = _workbook.Sheets[sheetName];
            ClearDataCells(sheet, S2_BLOCK1_START, S2_BLOCK1_END);
            ClearDataCells(sheet, S2_BLOCK2_START, S2_BLOCK2_END);
            SetMetaInt(sheetName, "blockCount", 1);
        }

        #endregion

        #region 1.3.3 단순 행 추가/삭제

        /// <summary>
        /// 시트의 특정 행 아래에 단순 행 추가. templateRow의 서식을 복사.
        /// </summary>
        public void AddSimpleRow(string sheetName, int headerRow, int firstDataRow)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var count = GetSimpleRowCount(sheetName, headerRow, firstDataRow);
            var insertRow = firstDataRow + count;

            dynamic templateRow = ws.Rows[firstDataRow];
            templateRow.Copy();
            ws.Rows[insertRow].Insert();
            dynamic destRow = ws.Rows[insertRow];
            destRow.PasteSpecial(-4122); // xlPasteFormats
            // 값 초기화 (B~R = 2~18)
            for (int c = 2; c <= 18; c++)
            {
                try { ws.Cells[insertRow, c].ClearContents(); } catch { }
            }
            _app.CutCopyMode = false;
        }

        public bool RemoveSimpleRow(string sheetName, int headerRow, int firstDataRow)
        {
            var count = GetSimpleRowCount(sheetName, headerRow, firstDataRow);
            if (count <= 1) return false;

            dynamic ws = _workbook.Sheets[sheetName];
            var lastRow = firstDataRow + count - 1;

            _app.DisplayAlerts = false;
            try { ws.Rows[lastRow].Delete(); }
            finally { _app.DisplayAlerts = true; }
            return true;
        }

        public int GetSimpleRowCount(string sheetName, int headerRow, int firstDataRow)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var count = 0;
            for (int r = firstDataRow; r <= firstDataRow + 500; r++)
            {
                // 테두리 또는 값이 있으면 카운트
                bool hasValue = false;
                for (int c = 2; c <= 18; c++)
                {
                    string v = ws.Cells[r, c].Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v)) { hasValue = true; break; }
                }

                bool hasBorder = false;
                if (!hasValue)
                {
                    try
                    {
                        dynamic borders = ws.Cells[r, 2].Borders;
                        hasBorder = borders[9].LineStyle != -4142; // xlNone
                    }
                    catch { }
                }

                if (!hasValue && !hasBorder) break;
                count++;
            }
            return count;
        }

        /// <summary>
        /// 메타 blockCount 기반 단순 행 추가. firstDataRow의 서식 복사.
        /// </summary>
        public void AddSimpleRowByMeta(string sheetName, int firstDataRow)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var count = GetMetaInt(sheetName, "blockCount", 1);
            var insertRow = firstDataRow + count;

            dynamic templateRow = ws.Rows[firstDataRow];
            templateRow.Copy();
            ws.Rows[insertRow].Insert();
            dynamic destRow = ws.Rows[insertRow];
            destRow.PasteSpecial(-4122); // xlPasteFormats
            for (int c = 2; c <= 18; c++)
            {
                try { ws.Cells[insertRow, c].ClearContents(); } catch { }
            }
            _app.CutCopyMode = false;

            SetMetaInt(sheetName, "blockCount", count + 1);
        }

        public bool RemoveSimpleRowByMeta(string sheetName, int firstDataRow)
        {
            var count = GetMetaInt(sheetName, "blockCount", 1);
            if (count <= 1) return false;

            dynamic ws = _workbook.Sheets[sheetName];
            var lastRow = firstDataRow + count - 1;

            _app.DisplayAlerts = false;
            try { ws.Rows[lastRow].Delete(); }
            finally { _app.DisplayAlerts = true; }

            SetMetaInt(sheetName, "blockCount", count - 1);
            return true;
        }

        #endregion

        #region 메타 시트 관리

        private void EnsureMetaSheet()
        {
            if (GetMetaSheet() != null) return;

            dynamic lastSheet = _workbook.Sheets[_workbook.Sheets.Count];
            dynamic newSheet = _workbook.Sheets.Add(After: lastSheet);
            newSheet.Name = MetaSheetName;
            newSheet.Visible = -1; // xlSheetVeryHidden

            newSheet.Cells[1, 1] = "key";
            newSheet.Cells[1, 2] = "value";

            var row = 2;

            // 시트 이름 기반 매핑 초기값
            var sheetMap = new (string section, string sheetName)[]
            {
                ("1.1~1.2", "1.1~1.2"),
                ("1.3.1",   "1.3.1"),
                ("1.3.2.1", "1.3.2.1"),
                ("1.3.2.2", "1.3.2.2"),
                ("1.3.3",   "1.3.3"),
                ("1.4",     "1.4"),
                ("2",       "2"),
                ("3.1~3.2.3.2", "3.1~3.2.3.2"),
            };

            foreach (var (section, name) in sheetMap)
            {
                // 시트가 실제로 존재하는지 확인
                bool exists = false;
                try { var _ = _workbook.Sheets[name]; exists = true; } catch { }
                if (exists)
                {
                    newSheet.Cells[row, 1] = $"sheet:{section}";
                    newSheet.Cells[row, 2] = name;
                    row++;
                }
            }

            // 행 블록 카운트 초기값 (1.3.1, 1.3.2.1, 1.3.2.2)
            var blockSheets = new[] { "1.3.1", "1.3.2.1", "1.3.2.2", "1.3.3", "1.4", "2", "3.1~3.2.3.2" };
            foreach (var name in blockSheets)
            {
                bool exists = false;
                try { var _ = _workbook.Sheets[name]; exists = true; } catch { }
                if (exists)
                {
                    newSheet.Cells[row, 1] = $"blockCount:{name}";
                    newSheet.Cells[row, 2] = 1;
                    row++;
                }
            }
        }

        private dynamic GetMetaSheet()
        {
            try { return _workbook.Sheets[MetaSheetName]; }
            catch { return null; }
        }

        private void AddMetaEntry(string section, string sheetName)
        {
            dynamic meta = GetMetaSheet();
            if (meta == null) return;
            var row = FindMetaEmptyRow(meta);
            meta.Cells[row, 1] = $"sheet:{section}";
            meta.Cells[row, 2] = sheetName;
        }

        private void RemoveMetaEntry(string section, string sheetName)
        {
            dynamic meta = GetMetaSheet();
            if (meta == null) return;
            var key = $"sheet:{section}";
            var row = 2;
            while (true)
            {
                string k = meta.Cells[row, 1].Value?.ToString();
                if (string.IsNullOrEmpty(k)) break;
                string v = meta.Cells[row, 2].Value?.ToString();
                if (k == key && v == sheetName) { meta.Rows[row].Delete(); return; }
                row++;
            }
        }

        private int GetMetaInt(string context, string key, int defaultValue)
        {
            dynamic meta = GetMetaSheet();
            if (meta == null) return defaultValue;
            var fullKey = $"{key}:{context}";
            var row = 2;
            while (true)
            {
                string k = meta.Cells[row, 1].Value?.ToString();
                if (string.IsNullOrEmpty(k)) break;
                if (k == fullKey)
                {
                    var val = meta.Cells[row, 2].Value;
                    return val != null ? Convert.ToInt32(val) : defaultValue;
                }
                row++;
            }
            return defaultValue;
        }

        private void SetMetaInt(string context, string key, int value)
        {
            dynamic meta = GetMetaSheet();
            if (meta == null) return;
            var fullKey = $"{key}:{context}";
            var row = 2;
            while (true)
            {
                string k = meta.Cells[row, 1].Value?.ToString();
                if (string.IsNullOrEmpty(k)) break;
                if (k == fullKey)
                {
                    meta.Cells[row, 2] = value;
                    return;
                }
                row++;
            }
            // 새 항목 추가
            row = FindMetaEmptyRow(meta);
            meta.Cells[row, 1] = fullKey;
            meta.Cells[row, 2] = value;
        }

        private int FindMetaEmptyRow(dynamic meta)
        {
            var row = 2;
            while (!string.IsNullOrEmpty(meta.Cells[row, 1].Value?.ToString()))
                row++;
            return row;
        }

        #endregion

        #region MappingOrchestrator용 메타 읽기

        /// <summary>
        /// _META에서 섹션→시트 매핑 목록 반환 (ClosedXML에서도 호출 가능하도록 static)
        /// </summary>
        public static List<(string section, string sheetName)> ReadSheetMappings(ClosedXML.Excel.IXLWorksheet metaWs)
        {
            var result = new List<(string, string)>();
            var row = 2;
            while (true)
            {
                var key = metaWs.Cell(row, 1).GetString()?.Trim();
                if (string.IsNullOrEmpty(key)) break;
                if (key.StartsWith("sheet:"))
                {
                    var section = key.Substring(6);
                    var sheetName = metaWs.Cell(row, 2).GetString()?.Trim();
                    result.Add((section, sheetName));
                }
                row++;
            }
            return result;
        }

        public static int ReadBlockCount(ClosedXML.Excel.IXLWorksheet metaWs, string sheetName)
        {
            var key = $"blockCount:{sheetName}";
            var row = 2;
            while (true)
            {
                var k = metaWs.Cell(row, 1).GetString()?.Trim();
                if (string.IsNullOrEmpty(k)) break;
                if (k == key)
                {
                    var val = metaWs.Cell(row, 2).GetString()?.Trim();
                    return int.TryParse(val, out var n) ? n : 1;
                }
                row++;
            }
            return 1;
        }

        #endregion

        #region Dispose

        private void QuitApp()
        {
            try { _app?.Quit(); } catch { }
            finally
            {
                if (_app != null) { Marshal.ReleaseComObject(_app); _app = null; }
                if (_workbook != null) { Marshal.ReleaseComObject(_workbook); _workbook = null; }
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            CloseWithSavePrompt();
        }

        #endregion
    }
}
