using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 1.4 글로벌최저한세 정보 요약 — Summary Collection에 매핑.
    /// 컬럼 헤더("1. 소재지국") 다음 행부터 데이터로 iteration.
    /// </summary>
    public class Mapping_1_4 : MappingBase
    {
        private const string COLUMN_HEADER_ANCHOR = "1. 소재지국"; // B열 컬럼헤더

        public Mapping_1_4()
            : base(null) { }

        public override void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        )
        {
            var headerRow = FindAnchorRow(ws, COLUMN_HEADER_ANCHOR);
            if (headerRow < 0) return;
            var dataStartRow = headerRow + 1;
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? dataStartRow;

            for (int row = dataStartRow; row <= lastRow; row++)
            {
                var summary = new Globe.GlobeBodyTypeSummary();
                summary.Jurisdiction = new Globe.SummaryTypeJurisdiction();

                // B: 소재지국
                var jurCode = ws.Cell(row, 2).GetString()?.Trim();
                if (!string.IsNullOrEmpty(jurCode))
                {
                    SetEnum<Globe.CountryCodeType>(
                        jurCode,
                        v =>
                        {
                            summary.Jurisdiction.JurisdictionName = v;
                            summary.Jurisdiction.JurisdictionNameSpecified = true;
                            // SummaryType.RecJurCode [R] — 행의 소재지국을 receiving 국가로 추가
                            summary.RecJurCode.Add(v);
                        },
                        errors,
                        fileName,
                        new MappingEntry { Cell = $"B{row}", Label = "소재지국" }
                    );
                }

                // C: 하위그룹 유형 / E: 하위그룹 최상위기업 TIN → Subgroup
                var subgroupTypeRaw = ws.Cell(row, 3).GetString()?.Trim();
                var subgroupTinRaw = ws.Cell(row, 5).GetString()?.Trim();
                if (!string.IsNullOrEmpty(subgroupTypeRaw) || !string.IsNullOrEmpty(subgroupTinRaw))
                {
                    var sg = new Globe.SummaryTypeJurisdictionSubgroup();
                    if (!string.IsNullOrEmpty(subgroupTypeRaw))
                    {
                        foreach (var code in subgroupTypeRaw.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
                            SetEnum<Globe.TypeofSubGroupEnumType>(
                                code,
                                v => sg.TypeofSubGroup.Add(v),
                                errors,
                                fileName,
                                new MappingEntry { Cell = $"C{row}", Label = "하위그룹 유형" }
                            );
                    }
                    if (!string.IsNullOrEmpty(subgroupTinRaw))
                        sg.Tin = ParseTin(subgroupTinRaw);
                    summary.Jurisdiction.Subgroup.Add(sg);
                }

                // G: 과세권 보유 국가 → Summary.JurWithTaxingRights
                var taxJurRaw = ws.Cell(row, 7).GetString()?.Trim();
                if (!string.IsNullOrEmpty(taxJurRaw))
                {
                    foreach (
                        var code in taxJurRaw.Split(
                            ',',
                            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
                        )
                    )
                    {
                        var jwr = new Globe.SummaryTypeJurWithTaxingRights();
                        SetEnum<Globe.CountryCodeType>(
                            code,
                            v =>
                            {
                                jwr.JurisdictionName = v;
                                jwr.JurisdictionNameSpecified = true;
                            },
                            errors,
                            fileName,
                            new MappingEntry { Cell = $"G{row}", Label = "과세권보유국가" }
                        );
                        if (jwr.JurisdictionNameSpecified)
                            summary.JurWithTaxingRights.Add(jwr);
                    }
                }

                // I-J: 적용면제/제외 사유 (SafeHarbour — Collection, 콤마 구분 다중값)
                var safeHarbour = ws.Cell(row, 9).GetString()?.Trim();
                if (!string.IsNullOrEmpty(safeHarbour))
                {
                    foreach (
                        var code in safeHarbour.Split(
                            ',',
                            StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries
                        )
                    )
                        SetEnum<Globe.SafeHarbourEnumType>(
                            code,
                            v => summary.SafeHarbour.Add(v),
                            errors,
                            fileName,
                            new MappingEntry { Cell = $"I{row}", Label = "적용면제" }
                        );
                }

                // K-L: 실효세율 범위
                var etrRange = ws.Cell(row, 11).GetString()?.Trim();
                if (!string.IsNullOrEmpty(etrRange))
                {
                    SetEnum<Globe.EtrRangeEnumType>(
                        etrRange,
                        v =>
                        {
                            summary.EtrRange = v;
                            summary.EtrRangeSpecified = true;
                        },
                        errors,
                        fileName,
                        new MappingEntry { Cell = $"K{row}", Label = "실효세율범위" }
                    );
                }

                // M-N: 실질기반제외소득 적용결과 추가세액 발생여부 → Sbie.NoTut
                //   "여" (발생) → NoTut=false,  "부" (미발생) → NoTut=true
                //   NotApplicable 은 별도 입력 없음 → false 고정 (SBIE 적용 대상)
                var sbieRaw = ws.Cell(row, 13).GetString()?.Trim();
                if (!string.IsNullOrEmpty(sbieRaw))
                {
                    summary.Sbie = new Globe.SummaryTypeSbie
                    {
                        NotApplicable = false,
                        NoTut = !ParseBool(sbieRaw),
                    };
                }

                // O-P: 추가세액(QDMTT) 범위
                var qdmtt = ws.Cell(row, 15).GetString()?.Trim();
                if (!string.IsNullOrEmpty(qdmtt))
                {
                    SetEnum<Globe.QdmtTuTEnumType>(
                        qdmtt,
                        v =>
                        {
                            summary.QdmtTut = v;
                            summary.QdmtTutSpecified = true;
                        },
                        errors,
                        fileName,
                        new MappingEntry { Cell = $"O{row}", Label = "QDMTT범위" }
                    );
                }

                // Q-R: 추가세액(GloBE) 범위
                var globeTut = ws.Cell(row, 17).GetString()?.Trim();
                if (!string.IsNullOrEmpty(globeTut))
                {
                    SetEnum<Globe.GlobeTuTEnumType>(
                        globeTut,
                        v =>
                        {
                            summary.GLoBeTut = v;
                            summary.GLoBeTutSpecified = true;
                        },
                        errors,
                        fileName,
                        new MappingEntry { Cell = $"Q{row}", Label = "GloBE범위" }
                    );
                }

                // 값이 하나라도 있으면 추가
                if (
                    !string.IsNullOrEmpty(jurCode)
                    || !string.IsNullOrEmpty(safeHarbour)
                    || !string.IsNullOrEmpty(etrRange)
                )
                    globe.GlobeBody.Summary.Add(summary);
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
