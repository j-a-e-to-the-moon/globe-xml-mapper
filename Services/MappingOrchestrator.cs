using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class MappingOrchestrator
    {
        // 섹션키 → 매퍼 생성 팩토리
        private static readonly Dictionary<string, Func<MappingBase>> MapperFactory = new()
        {
            { "1.1~1.2",  () => new Mapping_1_1_1_2() },
            { "1.3.1",    () => new Mapping_1_3_1() },
            { "1.3.2.1",  () => new Mapping_1_3_2_1() },
            { "1.3.2.2",  () => new Mapping_1_3_2_2() },
            // 1.3.3은 XML에 포함하지 않음 (AdditionalDataPoint 미사용)
            { "1.4",      () => new Mapping_1_4() },
            { "2",        () => new Mapping_2() },
        };

        /// <summary>
        /// 단일 Workbook + _META 숨김시트 기반 매핑.
        /// ControlPanelForm에서 호출.
        /// </summary>
        public List<string> MapWorkbook(string filePath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            using var workbook = new XLWorkbook(filePath);

            // _META 시트에서 섹션-시트 매핑 읽기
            if (!workbook.TryGetWorksheet(ExcelController.MetaSheetName, out var metaWs))
            {
                errors.Add("_META 숨김시트를 찾을 수 없습니다. Control Panel에서 파일을 먼저 열어주세요.");
                return errors;
            }

            var mappings = ExcelController.ReadSheetMappings(metaWs);

            foreach (var (section, sheetName) in mappings)
            {
                if (!MapperFactory.TryGetValue(section, out var createMapper))
                    continue; // 매퍼 없는 섹션은 스킵 (1.3.3 등 XML 미포함)

                if (!workbook.TryGetWorksheet(sheetName, out var ws))
                {
                    errors.Add($"시트 '{sheetName}'을(를) 찾을 수 없습니다. (섹션: {section})");
                    continue;
                }

                var mapper = createMapper();
                mapper.Map(ws, globe, errors, sheetName);
            }

            FillMessageSpec(globe);
            return errors;
        }

        /// <summary>
        /// 프로젝트 폴더 기반 매핑.
        /// 구조: root/main.xlsx + root/{n}/Group.xlsx + root/{n}/CE_*.xlsx
        /// </summary>
        public List<string> MapFolder(string rootPath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            // ── main.xlsx: _META 기반 매핑 (1.x, 2, 3.4.3 섹션) ──────────
            var mainPath = Path.Combine(rootPath, "main.xlsx");
            if (!File.Exists(mainPath))
            {
                errors.Add($"main.xlsx를 찾을 수 없습니다. ({rootPath})");
                return errors;
            }
            var mainErrors = MapWorkbook(mainPath, globe);
            errors.AddRange(mainErrors);

            // ── 숫자 폴더: Group.xlsx + CE_*.xlsx ──────────────────────────
            var countryDirs = Directory.GetDirectories(rootPath)
                .Where(d => int.TryParse(Path.GetFileName(d), out _))
                .OrderBy(d => int.Parse(Path.GetFileName(d)))
                .ToList();

            foreach (var dir in countryDirs)
            {
                var dirName = Path.GetFileName(dir);

                // Group.xlsx
                var groupPath = Path.Combine(dir, "Group.xlsx");
                if (File.Exists(groupPath))
                    MapFileBySheets(groupPath, globe, errors);
                else
                    errors.Add($"[{dirName}] Group.xlsx 없음");

                // CE_N.xlsx
                var ceFiles = Directory.GetFiles(dir, "CE_*.xlsx")
                    .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                    .OrderBy(f => f)
                    .ToList();
                foreach (var ceFile in ceFiles)
                    MapFileBySheets(ceFile, globe, errors);
            }

            FillMessageSpec(globe);
            return errors;
        }

        /// <summary>
        /// 파일 내 시트 이름을 MapperFactory에 직접 조회하여 매핑.
        /// _META 없이 동작 (Group.xlsx, CE_N.xlsx용).
        /// </summary>
        private void MapFileBySheets(string filePath, Globe.GlobeOecd globe, List<string> errors)
        {
            var fileName = Path.GetFileName(filePath);
            try
            {
                using var workbook = new XLWorkbook(filePath);
                foreach (var ws in workbook.Worksheets)
                {
                    if (!MapperFactory.TryGetValue(ws.Name, out var createMapper))
                        continue; // 아직 매퍼 없는 시트는 스킵
                    createMapper().Map(ws, globe, errors, fileName);
                }
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 파일 읽기 오류: {ex.Message}");
            }
        }

        #region 내부 유틸



        private static void ProcessFile(string filePath, MappingBase mapper, Globe.GlobeOecd globe, List<string> errors)
        {
            var fileName = Path.GetFileName(filePath);
            try
            {
                using var workbook = new XLWorkbook(filePath);
                foreach (var ws in workbook.Worksheets)
                {
                    if (ws.Name == mapper.SheetName)
                    {
                        mapper.Map(ws, globe, errors, fileName);
                        return;
                    }
                }
                errors.Add($"[{fileName}] 시트 '{mapper.SheetName}' 없음");
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 파일 읽기 오류: {ex.Message}");
            }
        }

        private static List<string> GetXlsxFiles(string dirPath)
        {
            return Directory.GetFiles(dirPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .OrderBy(f => f)
                .ToList();
        }

        #endregion

        #region MessageSpec / DocSpec

        private void FillMessageSpec(Globe.GlobeOecd globe)
        {
            var spec = globe.MessageSpec;
            var fi = globe.GlobeBody?.FilingInfo;

            if (fi?.FilingCe != null)
                spec.TransmittingCountry = fi.FilingCe.ResCountryCode;

            spec.ReceivingCountry = spec.TransmittingCountry;
            spec.MessageType = Globe.MessageTypeEnumType.Gir;

            if (fi?.Period != null && fi.Period.End != default)
                spec.ReportingPeriod = fi.Period.End;

            spec.Timestamp = DateTime.Now;

            if (string.IsNullOrEmpty(spec.MessageRefId))
            {
                var sendCC = spec.TransmittingCountry.ToString().ToUpper();
                var recvCC = spec.ReceivingCountry.ToString().ToUpper();
                var uid = spec.Timestamp.ToString("yyyyMMddHHmmss");
                spec.MessageRefId = $"{sendCC}{spec.ReportingPeriod:yyyy}{recvCC}{uid}";
            }

            FillDocSpecs(globe);
        }

        private void FillDocSpecs(Globe.GlobeOecd globe)
        {
            var sendCC = globe.MessageSpec.TransmittingCountry.ToString().ToUpper();
            var year = globe.MessageSpec.ReportingPeriod.ToString("yyyy");
            var ts = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            if (globe.GlobeBody.FilingInfo != null)
            {
                globe.GlobeBody.FilingInfo.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}FI{ts}"
                };
            }

            if (globe.GlobeBody.GeneralSection != null)
            {
                globe.GlobeBody.GeneralSection.DocSpec = new Globe.DocSpecType
                {
                    DocTypeIndic = Globe.OecdDocTypeIndicEnumType.Oecd1,
                    DocRefId = $"{sendCC}{year}GS{ts}"
                };
            }
        }

        #endregion
    }
}
