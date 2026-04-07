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
            try { _workbook.Close(SaveChanges: true); }
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

        #region CE 블록 + 별첨 시트 연동

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
            ws.Cells[refRow, 15] = $"별첨{count}"; // O열 = 15

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
            var target = $"별첨{attachNum}";
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
            for (int r = dataStart; r <= dataStart + 100; r++)
            {
                string val = ws.Cells[r, 2].Value?.ToString()?.Trim();
                // 빈 행이거나 다음 별첨 제목이면 종료
                if (string.IsNullOrEmpty(val))
                {
                    // 다음 행도 빈지 확인 (별첨 간 구분 빈행)
                    break;
                }
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

            // 행 삽입
            ws.Rows[insertRow].Insert();
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

            // 마지막 사용 행 찾기
            int lastRow = (int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count;

            var startRow = lastRow + 1;
            ws.Cells[startRow, 2] = $"별첨{attachNum}";
            ws.Cells[startRow + 2, 2] = "유형";
            ws.Cells[startRow + 2, 3] = "납세자번호";
            ws.Cells[startRow + 2, 4] = "소유지분(%)";
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
                if (val != null && val.StartsWith("별첨") && val != $"별첨{attachNum}")
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
            // 시트 매핑 초기값
            string sheet1 = null;
            for (int i = 1; i <= _workbook.Sheets.Count; i++)
            {
                string name = _workbook.Sheets[i].Name;
                if (name.Contains("(1)")) { sheet1 = name; break; }
            }

            if (sheet1 != null)
            {
                newSheet.Cells[row, 1] = "sheet:1.1~1.2"; newSheet.Cells[row, 2] = sheet1; row++;
                newSheet.Cells[row, 1] = "sheet:1.3.1";   newSheet.Cells[row, 2] = sheet1; row++;
            }

            for (int i = 1; i <= _workbook.Sheets.Count; i++)
            {
                string name = _workbook.Sheets[i].Name;
                if (name == MetaSheetName || name == sheet1) continue;
                if (name.Contains("(2)"))
                { newSheet.Cells[row, 1] = "sheet:1.3.2.1"; newSheet.Cells[row, 2] = name; row++; }
                else if (name.Contains("(3)"))
                { newSheet.Cells[row, 1] = "sheet:1.3.2.2"; newSheet.Cells[row, 2] = name; row++; }
            }

            // 블록 카운트 초기값
            if (sheet1 != null)
            {
                newSheet.Cells[row, 1] = $"blockCount:{sheet1}"; newSheet.Cells[row, 2] = 1; row++;
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
