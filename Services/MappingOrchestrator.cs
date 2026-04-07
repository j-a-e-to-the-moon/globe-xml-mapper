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
                {
                    errors.Add($"알 수 없는 섹션 '{section}' (시트: {sheetName})");
                    continue;
                }

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
        /// 디렉토리 기반 매핑 (하위 호환용).
        /// </summary>
        public List<string> MapFolder(string rootPath, Globe.GlobeOecd globe)
        {
            var errors = new List<string>();
            var filingMapper = new Mapping_1_1_1_2();

            // 루트: 기본정보
            var rootFiles = GetXlsxFiles(rootPath);
            foreach (var f in rootFiles)
                ProcessFile(f, filingMapper, globe, errors);

            // 하위 디렉토리
            var subDirs = new[] { ("1.3.1", (Func<MappingBase>)(() => new Mapping_1_3_1())),
                                  ("1.3.2.1", () => new Mapping_1_3_2_1()),
                                  ("1.3.2.2", () => new Mapping_1_3_2_2()) };

            foreach (var (dirName, createMapper) in subDirs)
            {
                var subDir = Path.Combine(rootPath, dirName);
                if (!Directory.Exists(subDir)) { errors.Add($"필수 디렉토리 '{dirName}' 없음"); continue; }
                var files = GetXlsxFiles(subDir);
                if (files.Count == 0) { errors.Add($"'{dirName}' 디렉토리에 xlsx 없음"); continue; }
                foreach (var f in files)
                {
                    var mapper = createMapper();
                    ProcessFile(f, mapper, globe, errors);
                }
            }

            FillMessageSpec(globe);
            return errors;
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
