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
        private bool _ownsApp;   // Open()으로 직접 생성한 경우만 true → Quit() 호출 여부 결정

        public const string MetaSheetName = "_META";

        public event Action WorkbookClosed;

        public bool IsOpen
        {
            get
            {
                // _app.Visible 대신 _workbook.Name 접근으로 실제 열림 여부 확인.
                // Excel이 다이얼로그/작업 중일 때 Visible은 false가 아니지만 COM이 throw할 수 있음.
                try
                {
                    if (_workbook == null || _app == null) return false;
                    var _ = (string)_workbook.Name;
                    return true;
                }
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
            _ownsApp = true;

            dynamic workbooks = _app.Workbooks;
            _workbook = workbooks.Open(path);
            Marshal.ReleaseComObject(workbooks);

            dynamic firstSheet = _workbook.Sheets[1];
            EnsureMetaSheet();
            firstSheet.Activate();
            Marshal.ReleaseComObject(firstSheet);
        }

        // oleaut32.dll의 GetActiveObject 직접 호출 (Marshal.GetActiveObject는 .NET Core에서 제거됨)
        [System.Runtime.InteropServices.DllImport("oleaut32.dll")]
        private static extern void GetActiveObjectNative(
            ref Guid rclsid,
            IntPtr pvReserved,
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IUnknown)]
            out object ppunk);

        // Excel.Application CLSID (버전 무관 공통)
        private static readonly Guid ExcelClsid = new Guid("00024500-0000-0000-C000-000000000046");

        /// <summary>
        /// 이미 열려 있는 Excel 인스턴스의 활성 워크북에 연결.
        /// 사용자가 직접 Excel에서 파일을 열고 이 메서드를 호출.
        /// </summary>
        public void AttachToActive()
        {
            try
            {
                var clsid = ExcelClsid;
                GetActiveObjectNative(ref clsid, IntPtr.Zero, out object obj);
                _app = obj;
            }
            catch (Exception ex) when (!(ex is InvalidOperationException))
            {
                throw new InvalidOperationException(
                    "실행 중인 Excel 인스턴스를 찾을 수 없습니다.\nExcel에서 파일을 먼저 열어주세요.", ex);
            }

            _workbook = _app.ActiveWorkbook;
            if (_workbook == null)
                throw new InvalidOperationException(
                    "열려 있는 Excel 통합문서가 없습니다.\nExcel에서 파일을 열고 다시 시도하세요.");

            _ownsApp = false;  // 외부 인스턴스에 연결 → Quit() 하지 않음
            EnsureMetaSheet();
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
            bool doQuit = false;
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

                    if (result == System.Windows.Forms.DialogResult.Cancel) return; // Quit 하지 않음
                    _workbook.Close(SaveChanges: result == System.Windows.Forms.DialogResult.Yes);
                }
                else
                {
                    _workbook.Close(SaveChanges: false);
                }
                doQuit = true;
            }
            catch { doQuit = true; }
            finally { if (doQuit) QuitApp(); }
        }

        /// <summary>
        /// COM RPC_E_CALL_REJECTED(0x80010001) 발생 시 최대 maxRetries회 재시도.
        /// Excel이 일시적으로 바쁜 상태일 때 COM 호출이 거부되는 것을 처리.
        /// </summary>
        private static void ComRetry(Action action, int maxRetries = 5, int delayMs = 150)
        {
            for (int i = 0; ; i++)
            {
                try { action(); return; }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80010001) && i < maxRetries)
                {
                    System.Threading.Thread.Sleep(delayMs);
                }
            }
        }

        #region 시트 내 행 블록 반복 (1.3.1 UPE)

        /// <summary>
        /// B열에서 header 텍스트를 포함하는 행 번호 목록 반환.
        /// blockHeader 기반 탐지에 사용.
        /// </summary>
        private List<int> FindBlockHeaderRows(dynamic ws, string header)
        {
            var result = new List<int>();
            int lastRow;
            try { lastRow = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count; }
            catch { lastRow = 500; }

            for (int r = 1; r <= lastRow; r++)
            {
                string val = ws.Cells[r, 2].Value?.ToString()?.Trim();
                if (val != null && val.Contains(header))
                    result.Add(r);
            }
            return result;
        }

        /// <summary>
        /// 시트 내 행 블록을 복제하여 아래에 추가.
        /// blockHeader 지정 시 B열 헤더 스캔으로 마지막 블록 위치를 동적 탐지.
        /// blockHeader 미지정 시 _META blockCount 기반.
        /// </summary>
        public void AddRowBlock(string sheetName, int sourceStartRow, int sourceEndRow, int gap,
            int dataColStart = 15, int dataColEnd = 18, string blockHeader = null)
        {
            ComRetry(() =>
            {
            dynamic ws = _workbook.Sheets[sheetName];
            var blockSize = sourceEndRow - sourceStartRow + 1;

            int insertRow;
            int count;

            if (blockHeader != null)
            {
                var headerRows = FindBlockHeaderRows(ws, blockHeader);
                count = headerRows.Count;
                var lastStart = count > 0 ? headerRows[count - 1] : sourceStartRow;
                var lastEnd = lastStart + (sourceEndRow - sourceStartRow);
                insertRow = lastEnd + 1 + gap;
            }
            else
            {
                count = GetMetaInt(sheetName, "blockCount", 1);
                insertRow = sourceEndRow + 1 + (count - 1) * (blockSize + gap) + gap;
            }

            // 빈 행 삽입
            dynamic insertRange = ws.Rows[$"{insertRow}:{insertRow + blockSize - 1}"];
            insertRange.Insert();

            // 원본 블록 복사
            dynamic sourceRange = ws.Range[
                ws.Cells[sourceStartRow, 1],
                ws.Cells[sourceEndRow, 18]
            ];
            dynamic destRange = ws.Range[
                ws.Cells[insertRow, 1],
                ws.Cells[insertRow + blockSize - 1, 18]
            ];
            sourceRange.Copy(destRange);

            // 행 높이 복사
            for (int i = 0; i < blockSize; i++)
            {
                try
                {
                    var srcH = ws.Rows[sourceStartRow + i].RowHeight;
                    if (srcH != null)
                        ws.Rows[insertRow + i].RowHeight = (double)srcH;
                }
                catch { }
            }

            // 데이터 셀 초기화 (값만 지우기, 서식 유지)
            ClearDataCells(ws, insertRow, insertRow + blockSize - 1, dataColStart, dataColEnd);

            if (blockHeader == null)
                SetMetaInt(sheetName, "blockCount", count + 1);
            }); // ComRetry
        }

        /// <summary>
        /// 마지막 행 블록 삭제.
        /// blockHeader 지정 시 B열 헤더 스캔으로 마지막 블록 위치를 동적 탐지.
        /// </summary>
        public bool RemoveRowBlock(string sheetName, int sourceStartRow, int sourceEndRow, int gap,
            string blockHeader = null)
        {
            bool result = false;
            ComRetry(() =>
            {
            dynamic ws = _workbook.Sheets[sheetName];
            var blockSize = sourceEndRow - sourceStartRow + 1;

            int count;
            int lastBlockStart, lastBlockEnd;

            if (blockHeader != null)
            {
                var headerRows = FindBlockHeaderRows(ws, blockHeader);
                count = headerRows.Count;
                if (count <= 1) return;
                lastBlockStart = headerRows[count - 1];
                lastBlockEnd = lastBlockStart + (sourceEndRow - sourceStartRow);
            }
            else
            {
                count = GetMetaInt(sheetName, "blockCount", 1);
                if (count <= 1) return;
                lastBlockStart = sourceEndRow + 1 + (count - 2) * (blockSize + gap) + gap;
                lastBlockEnd = lastBlockStart + blockSize - 1;
            }

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

            if (blockHeader == null)
                SetMetaInt(sheetName, "blockCount", count - 1);
            result = true;
            }); // ComRetry
            return result;
        }

        /// <summary>
        /// 시트를 원래 상태로 초기화 (추가된 블록 모두 제거 + 데이터 초기화).
        /// blockHeader 지정 시 B열 헤더 스캔 사용.
        /// </summary>
        public void ResetSheet(string sheetName, int sourceStartRow, int sourceEndRow, int gap,
            int dataColStart = 15, int dataColEnd = 18, string blockHeader = null)
        {
            if (blockHeader != null)
            {
                // 헤더 탐색 없이 sourceEndRow+1 ~ lastUsed 전체 삭제.
                // 어떤 상태(블록이 내부에 삽입됐든, 헤더가 없든)에서도 확실히 초기화.
                dynamic ws = _workbook.Sheets[sheetName];
                int deleteFrom = sourceEndRow + 1;
                int lastUsed;
                try { lastUsed = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count - 1; }
                catch { lastUsed = deleteFrom + 300; }

                if (deleteFrom <= lastUsed)
                {
                    _app.DisplayAlerts = false;
                    try
                    {
                        dynamic deleteRange = ws.Rows[$"{deleteFrom}:{lastUsed}"];
                        deleteRange.Delete();
                    }
                    finally { _app.DisplayAlerts = true; }
                }
            }
            else
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
                    finally { _app.DisplayAlerts = true; }
                }
                SetMetaInt(sheetName, "blockCount", 1);
            }

            // 첫 번째 블록 데이터 초기화
            dynamic sheet = _workbook.Sheets[sheetName];
            ClearDataCells(sheet, sourceStartRow, sourceEndRow, dataColStart, dataColEnd);
        }

        /// <summary>
        /// blockHeader 지정 시 B열 헤더 스캔으로 블록 수 반환, 미지정 시 _META 사용.
        /// </summary>
        public int GetRowBlockCount(string sheetName, int defaultCount = 1, string blockHeader = null)
        {
            if (blockHeader != null)
            {
                dynamic ws = _workbook.Sheets[sheetName];
                return FindBlockHeaderRows(ws, blockHeader).Count;
            }
            return GetMetaInt(sheetName, "blockCount", defaultCount);
        }

        private void ClearDataCells(dynamic ws, int startRow, int endRow,
            int dataColStart = 15, int dataColEnd = 18)
        {
            for (int r = startRow; r <= endRow; r++)
            {
                for (int c = dataColStart; c <= dataColEnd; c++)
                {
                    dynamic cell = ws.Cells[r, c];
                    if (cell.MergeCells)
                    {
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

        private const int CE_BLOCK_START = 3;
        private const int CE_BLOCK_END = 21;  // row 3~21 = 19행 (기존 20은 오류)
        private const int CE_BLOCK_GAP = 2;
        private const int CE_ATTACH_REF_ROW_OFFSET = 8; // 블록 내 O11 = 시작행+8

        /// <summary>
        /// CE 블록 추가: 행 블록 복제(헤더 기반) + 첨부N 셀 갱신 + 별첨 시트 섹션 추가.
        /// </summary>
        public void AddCeBlock(string ceSheetName, string attachSheetName)
        {
            // 1. 헤더 기반으로 행 블록 복제
            AddRowBlock(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP,
                blockHeader: "1.3.2.1");

            // 2. 삽입 후 헤더 재스캔 → 새 블록 시작행 확인
            dynamic ws = _workbook.Sheets[ceSheetName];
            var headerRows = FindBlockHeaderRows(ws, "1.3.2.1");
            var count = headerRows.Count;
            var newBlockStart = headerRows[count - 1];
            ws.Cells[newBlockStart + CE_ATTACH_REF_ROW_OFFSET, 15] = $"첨부{count}";

            // 3. 별첨 시트에 섹션 추가
            AddAttachSection(attachSheetName, count);
        }

        /// <summary>
        /// 마지막 CE 블록 삭제 + 별첨 시트에서 해당 섹션 삭제.
        /// </summary>
        public bool RemoveCeBlock(string ceSheetName, string attachSheetName)
        {
            dynamic ws = _workbook.Sheets[ceSheetName];
            var count = FindBlockHeaderRows(ws, "1.3.2.1").Count;
            if (count <= 1) return false;

            RemoveRowBlock(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP,
                blockHeader: "1.3.2.1");
            RemoveAttachSection(attachSheetName, count);
            return true;
        }

        /// <summary>
        /// CE 시트 + 별첨 시트 초기화.
        /// </summary>
        public void ResetCeSheet(string ceSheetName, string attachSheetName)
        {
            dynamic ws = _workbook.Sheets[ceSheetName];
            var count = FindBlockHeaderRows(ws, "1.3.2.1").Count;

            if (count > 1)
                for (int i = count; i >= 2; i--)
                    RemoveAttachSection(attachSheetName, i);
            ResetAttachSection(attachSheetName, 1);

            ResetSheet(ceSheetName, CE_BLOCK_START, CE_BLOCK_END, CE_BLOCK_GAP,
                blockHeader: "1.3.2.1");
        }

        public int GetCeBlockCount(string ceSheetName)
        {
            dynamic ws = _workbook.Sheets[ceSheetName];
            return FindBlockHeaderRows(ws, "1.3.2.1").Count;
        }

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

            // 클립보드를 비운 후 Insert → 빈 행 삽입 (클립보드가 있으면 Insert가 붙여넣기를 같이 해버림)
            _app.CutCopyMode = false;
            ws.Rows[insertRow].Insert();

            // 삽입된 빈 행에 서식 복사
            ws.Rows[dataStart].Copy();
            ws.Rows[insertRow].PasteSpecial(-4122); // xlPasteFormats = -4122
            _app.CutCopyMode = false;

            // 값 초기화 (PasteFormats가 내용을 복사하지 않지만 안전하게 비움)
            ws.Cells[insertRow, 2].ClearContents();
            ws.Cells[insertRow, 3].ClearContents();
            ws.Cells[insertRow, 4].ClearContents();
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

            var attach1Start = FindAttachSectionStart(ws, 1);
            if (attach1Start < 0) return;

            // 마지막 사용 행 뒤에 1행 간격 두고 시작
            int lastRow  = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count;
            int startRow = lastRow + 1;

            // 첨부1의 title+empty+header+첫 데이터행(4행)을 통째로 복사 → 폰트/테두리/행높이 완전 보존
            int copyCount = ATTACH_HEADER_ROWS + ATTACH_INITIAL_DATA_ROWS; // 3+1=4
            _app.CutCopyMode = false;
            dynamic srcRange = ws.Range[ws.Cells[attach1Start, 1],
                                        ws.Cells[attach1Start + copyCount - 1, 10]];
            dynamic dstRange = ws.Range[ws.Cells[startRow, 1],
                                        ws.Cells[startRow + copyCount - 1, 10]];
            srcRange.Copy(dstRange);

            // Range.Copy는 행 높이를 복사하지 않으므로 명시적 복사
            for (int i = 0; i < copyCount; i++)
                ws.Rows[startRow + i].RowHeight = ws.Rows[attach1Start + i].RowHeight;

            // 제목 갱신 + 데이터 초기화
            ws.Cells[startRow, 2].Value = $"첨부{attachNum}";
            ws.Cells[startRow + ATTACH_HEADER_ROWS, 2].ClearContents();
            ws.Cells[startRow + ATTACH_HEADER_ROWS, 3].ClearContents();
            ws.Cells[startRow + ATTACH_HEADER_ROWS, 4].ClearContents();

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

        #region 시트3 대형 블록 (2~219, 페이지번호 행 삭제됨)

        // 페이지번호 행이 템플릿에서 삭제되었으므로 더 이상 사용하지 않음
        private const int S3_BLOCK_START = 2;
        private const int S3_BLOCK_END = 219;
        private const int S3_PAGE_GAP = 2; // 페이지 간 간격

        public void AddSheet3Page(string sheetName)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var count = GetMetaInt(sheetName, "blockCount", 1);
            var blockSize = S3_BLOCK_END - S3_BLOCK_START + 1; // 226행

            // 삽입 위치: 시트 끝에 직접 복사 (Insert 안 함 — 병합 셀 충돌 방지)
            var insertRow = S3_BLOCK_END + 1 + (count - 1) * (blockSize + S3_PAGE_GAP) + S3_PAGE_GAP;

            // 원본 전체를 시트 끝에 직접 복사
            dynamic srcRange = ws.Range[ws.Cells[S3_BLOCK_START, 1], ws.Cells[S3_BLOCK_END, 18]];
            dynamic dstRange = ws.Range[ws.Cells[insertRow, 1], ws.Cells[insertRow + blockSize - 1, 18]];
            srcRange.Copy(dstRange);

            // 행 높이 복사
            for (int i = 0; i < blockSize; i++)
                ws.Rows[insertRow + i].RowHeight = (double)ws.Rows[S3_BLOCK_START + i].RowHeight;

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
            // subKey 형태: "p1:cfc", "p2:carryback" 등
            var key = subKey.Contains(':') ? subKey.Split(':').Last() : subKey;
            return key switch
            {
                "cfc" => 2,       // 통합형피지배 초기 2행 (101~102)
                "carryback" => 2, // 결손금 소급공제 초기 2행 (145~146)
                "art89" => 5,     // 제89조 초기 5행 (176~180)
                _ => 1
            };
        }

        #endregion

        #region 시트2 복합 블록 (3~23 + 26~54)

        // 시트2는 블록1(2~22) + 간격(23~24) + 블록2(25~53) = 총 52행이 한 세트
        private const int S2_BLOCK1_START = 2;
        private const int S2_BLOCK1_END = 22;
        private const int S2_GAP_ROWS = 2;  // 23~24행 (간격)
        private const int S2_BLOCK2_START = 25;
        private const int S2_BLOCK2_END = 53;
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

        #region 적용면제 첨부 시트 관리

        private const string S2_ATTACH_SHEET = "적용면제 첨부";

        /// <summary>
        /// "적용면제 첨부" 시트에 새 첨부N 섹션 추가. 시트가 없으면 무시.
        /// </summary>
        public void AddSheet2AttachPage(int blockNum)
        {
            try { var _ = _workbook.Sheets[S2_ATTACH_SHEET]; }
            catch { return; } // 시트 없으면 스킵
            AddAttachSection(S2_ATTACH_SHEET, blockNum);
        }

        /// <summary>
        /// "적용면제 첨부" 시트에서 마지막 첨부 섹션 삭제. 시트가 없으면 무시.
        /// </summary>
        public void RemoveSheet2AttachPage(int blockNum)
        {
            try { var _ = _workbook.Sheets[S2_ATTACH_SHEET]; }
            catch { return; }
            RemoveAttachSection(S2_ATTACH_SHEET, blockNum);
        }

        /// <summary>
        /// "적용면제 첨부" 시트가 있으면 첨부 섹션 수 반환, 없으면 0.
        /// </summary>
        public int GetSheet2AttachPageCount()
        {
            dynamic ws;
            try { ws = _workbook.Sheets[S2_ATTACH_SHEET]; }
            catch { return 0; }

            // 첨부N 헤더 행 수 카운트
            int count = 0;
            for (int r = 1; r <= 5000; r++)
            {
                string val = ws.Cells[r, 2].Value?.ToString()?.Trim();
                if (string.IsNullOrEmpty(val) && r > 10) break;
                if (val != null && System.Text.RegularExpressions.Regex.IsMatch(val, @"^첨부\d+$"))
                    count++;
            }
            return count;
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
        public void AddSimpleRowByMeta(string sheetName, int firstDataRow, int defaultCount = 1)
        {
            dynamic ws = _workbook.Sheets[sheetName];
            var count = GetMetaInt(sheetName, "blockCount", defaultCount);
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

        public bool RemoveSimpleRowByMeta(string sheetName, int firstDataRow, int defaultCount = 1)
        {
            var count = GetMetaInt(sheetName, "blockCount", defaultCount);
            if (count <= defaultCount) return false;

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

        // main_template 섹션→시트 매핑 (ControlPanel + _META용)
        internal static readonly (string section, string sheetName)[] SheetMap = new[]
        {
            ("1.1~1.2",     "다국적기업그룹 정보"),
            ("1.3.1",       "최종모기업"),
            ("1.3.2.1",     "그룹구조"),
            ("1.3.2.2",     "제외기업"),
            ("1.3.3",       "그룹구조 변동"),
            ("1.4",         "요약"),
            ("2",           "적용면제"),
            ("UTPR",        "UTPR 배분"),
            ("JurCal",      "3.1~3.2.3.2"),
        };

        // fileType별 섹션→시트 매핑 (MapFileBySheets용, 코드로 관리)
        internal static readonly Dictionary<string, (string section, string sheetName)[]> FileTypeSheetMap = new()
        {
            ["group"] = new[]
            {
                ("JurCal", "국가별 계산"),
                ("2",      "추가세액 계산"),
            },
            ["entity"] = new[]
            {
                ("EntityCe", "구성기업 계산"),
            },
        };

        private void EnsureMetaSheet()
        {
            dynamic meta = GetMetaSheet();

            if (meta == null)
            {
                dynamic lastSheet = _workbook.Sheets[_workbook.Sheets.Count];
                meta = _workbook.Sheets.Add(After: lastSheet);
                meta.Name = MetaSheetName;
                meta.Visible = -1; // xlSheetVeryHidden
                meta.Cells[1, 1] = "key";
                meta.Cells[1, 2] = "value";
            }

            // 기존 blockCount / fileType 값 보존
            var savedBlockCounts = new Dictionary<string, int>();
            string savedFileType = null;
            var r = 2;
            while (true)
            {
                var k = (string)meta.Cells[r, 1].Value?.ToString();
                if (string.IsNullOrEmpty(k)) break;
                if (k.StartsWith("blockCount:"))
                {
                    var name = k.Substring(11);
                    if (int.TryParse((string)meta.Cells[r, 2].Value?.ToString(), out var cnt))
                        savedBlockCounts[name] = cnt;
                }
                else if (k == "fileType")
                    savedFileType = meta.Cells[r, 2].Value?.ToString();
                r++;
            }

            // 헤더 이후 전체 초기화
            for (var clearRow = 2; clearRow < r + 1; clearRow++)
            {
                meta.Cells[clearRow, 1] = "";
                meta.Cells[clearRow, 2] = "";
            }

            // sheet: 항목 재작성 (실제 존재하는 시트만)
            var row = 2;
            foreach (var (section, name) in SheetMap)
            {
                bool exists = false;
                try { var _ = _workbook.Sheets[name]; exists = true; } catch { }
                if (exists)
                {
                    meta.Cells[row, 1] = $"sheet:{section}";
                    meta.Cells[row, 2] = name;
                    row++;
                }
            }

            // blockCount: 항목 재작성 (보존된 값 또는 1)
            var blockSheets = new[] { "최종모기업", "그룹구조", "제외기업", "그룹구조 변동", "요약", "적용면제", "3.1~3.2.3.2" };
            foreach (var name in blockSheets)
            {
                bool exists = false;
                try { var _ = _workbook.Sheets[name]; exists = true; } catch { }
                if (exists)
                {
                    meta.Cells[row, 1] = $"blockCount:{name}";
                    meta.Cells[row, 2] = savedBlockCounts.TryGetValue(name, out var cnt) ? cnt : 1;
                    row++;
                }
            }

            // fileType: 보존된 값 유지 (없으면 "main" 기본값)
            meta.Cells[row, 1] = "fileType";
            meta.Cells[row, 2] = savedFileType ?? "main";
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

        /// <summary>
        /// _META에서 fileType 값 반환. "main" / "group" / "entity"
        /// </summary>
        public static string ReadFileType(ClosedXML.Excel.IXLWorksheet metaWs)
        {
            var row = 2;
            while (true)
            {
                var key = metaWs.Cell(row, 1).GetString()?.Trim();
                if (string.IsNullOrEmpty(key)) break;
                if (key == "fileType")
                    return metaWs.Cell(row, 2).GetString()?.Trim() ?? "main";
                row++;
            }
            return "main"; // 기본값
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
            // 순서 중요: 워크북 → (소유 시) Quit → 앱 → GC
            // GC.Collect 없이는 dynamic으로 생성된 중간 COM 객체(RCW)가 남아 Excel 프로세스가 살아있게 됨
            try
            {
                if (_workbook != null)
                {
                    Marshal.ReleaseComObject(_workbook);
                    _workbook = null;
                }
                // AttachToActive()로 연결한 경우 Quit() 호출 금지 — 사용자 Excel을 닫지 않음
                if (_ownsApp)
                    _app?.Quit();
            }
            catch { }
            finally
            {
                if (_app != null)
                {
                    Marshal.ReleaseComObject(_app);
                    _app = null;
                }
                _ownsApp = false;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
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
