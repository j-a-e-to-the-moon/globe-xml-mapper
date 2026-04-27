using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace GlobeMapper.Services
{
    /// <summary>
    /// GIR XML 에러코드 매뉴얼 기반 유효성 검증
    /// </summary>
    public static class ValidationUtil
    {
        public static List<string> Validate(Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            ValidateMessageSpec(globe, errors);
            ValidateDocSpec(globe, errors);
            ValidateFilingInfo(globe, errors);
            ValidateRecJurCode(globe, errors);
            ValidateUpe(globe, errors);
            ValidateCe(globe, errors);
            ValidateCeOwnershipTin(globe, errors);
            ValidateRulesConsistency(globe, errors);
            ValidateTin(globe, errors);
            ValidateSummary(globe, errors);
            ValidateSafeHarbourCompleteness(globe, errors);
            ValidateJurisdictionSection(globe, errors);
            ValidateUtprAttribution(globe, errors);
            ValidateDeminimisDataPresence(globe, errors);

            return errors;
        }

        /// <summary>
        /// [데이터경고] Deminimis 세이프하버(GIR2901/GIR2902) 선택 시 financial 데이터 모두 0이면 경고.
        /// XSD는 valid이지만 사용자가 CbCR 수치를 안 채웠을 가능성 큼.
        /// </summary>
        private static void ValidateDeminimisDataPresence(Globe.GlobeOecd globe, List<string> errors)
        {
            if (globe.GlobeBody?.JurisdictionSection == null) return;

            foreach (var js in globe.GlobeBody.JurisdictionSection)
            {
                var jur = js.Jurisdiction.ToString().ToUpper();
                foreach (var etr in js.GLoBeTax?.Etr ?? Enumerable.Empty<Globe.EtrType>())
                {
                    var dm = etr.EtrStatus?.EtrException?.DeminimisSimplifiedNmceCalc;
                    if (dm == null) continue;

                    bool hasNonZero = false;
                    foreach (var fd in dm.FinancialData)
                    {
                        if (TryParseDec(fd.Revenue, out var r) && r != 0) { hasNonZero = true; break; }
                        if (TryParseDec(fd.GlobeRevenue, out var gr) && gr != 0) { hasNonZero = true; break; }
                        if (TryParseDec(fd.NetGlobeIncome, out var ni) && ni != 0) { hasNonZero = true; break; }
                        if (TryParseDec(fd.Fanil, out var fa) && fa != 0) { hasNonZero = true; break; }
                    }
                    if (!hasNonZero && dm.Average != null)
                    {
                        if (TryParseDec(dm.Average.Revenue, out var ar) && ar != 0) hasNonZero = true;
                        else if (TryParseDec(dm.Average.GlobeRevenue, out var agr) && agr != 0) hasNonZero = true;
                        else if (TryParseDec(dm.Average.NetGlobeIncome, out var ani) && ani != 0) hasNonZero = true;
                        else if (TryParseDec(dm.Average.Fanil, out var afa) && afa != 0) hasNonZero = true;
                    }

                    if (!hasNonZero)
                        errors.Add(
                            $"[데이터경고] [{jur}] Deminimis 세이프하버({dm.Basis})을 선택했으나 FinancialData/Average 모두 0임. CbCR/회계 수치를 확인하세요."
                        );
                }
            }
        }

        #region MessageSpec (60001, 60003)

        private static void ValidateMessageSpec(Globe.GlobeOecd globe, List<string> errors)
        {
            var spec = globe.MessageSpec;
            if (spec == null)
                return;

            // 60001: MessageRefId 형식 [발신국가코드][보고기간][수신국가코드][고유식별번호]
            if (string.IsNullOrEmpty(spec.MessageRefId))
            {
                errors.Add("[60001] [MessageSpec] MessageRefId가 비어 있습니다.");
            }

            // 60003: ReportingPeriod YYYY ≤ 현재연도
            if (spec.ReportingPeriod != default && spec.ReportingPeriod.Year > DateTime.Now.Year)
            {
                errors.Add(
                    $"[60003] [MessageSpec] ReportingPeriod 연도({spec.ReportingPeriod.Year})가 현재 연도({DateTime.Now.Year})보다 큽니다."
                );
            }
        }

        #endregion

        #region DocSpec (60004, 60011, 60012, 60017)

        private static void ValidateDocSpec(Globe.GlobeOecd globe, List<string> errors)
        {
            var body = globe.GlobeBody;
            if (body == null)
                return;

            // 60016: OECD0(재제출) FilingInfo이면 GeneralSection은 OECD1 불가
            if (
                body.FilingInfo?.DocSpec?.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd0
                && body.GeneralSection?.DocSpec?.DocTypeIndic
                    == Globe.OecdDocTypeIndicEnumType.Oecd1
            )
            {
                errors.Add(
                    "[60016] [DocSpec] FilingInfo DocTypeIndic이 OECD0(재제출)인 경우 GeneralSection은 OECD1(신규)이 될 수 없습니다."
                );
            }

            // 60013: OECD0(재제출)은 FilingInfo에만 사용 가능
            if (body.GeneralSection?.DocSpec?.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd0)
                errors.Add(
                    "[60013] [GeneralSection/DocSpec] OECD0(재제출)은 FilingInfo에만 사용 가능합니다. GeneralSection에 OECD0을 사용할 수 없습니다."
                );

            // 60017: FilingInfo DocTypeIndic=OECD1이면 GeneralSection 필수
            if (
                body.FilingInfo?.DocSpec?.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd1
                && body.GeneralSection == null
            )
            {
                errors.Add(
                    "[60017] [FilingInfo/DocSpec] DocTypeIndic이 OECD1이면 GeneralSection이 필수입니다."
                );
            }

            // FilingInfo DocSpec 검증
            ValidateSingleDocSpec(body.FilingInfo?.DocSpec, "FilingInfo", errors);

            // GeneralSection DocSpec 검증
            ValidateSingleDocSpec(body.GeneralSection?.DocSpec, "GeneralSection", errors);

            // 60004: 신규(OECD1)와 정정(OECD2/OECD3) 혼합 불가
            var docTypes = new List<Globe.OecdDocTypeIndicEnumType?>();
            docTypes.Add(body.FilingInfo?.DocSpec?.DocTypeIndic);
            docTypes.Add(body.GeneralSection?.DocSpec?.DocTypeIndic);
            var distinct = docTypes.Where(d => d.HasValue).Select(d => d.Value).Distinct().ToList();
            if (
                distinct.Contains(Globe.OecdDocTypeIndicEnumType.Oecd1)
                && (
                    distinct.Contains(Globe.OecdDocTypeIndicEnumType.Oecd2)
                    || distinct.Contains(Globe.OecdDocTypeIndicEnumType.Oecd3)
                )
            )
            {
                errors.Add(
                    "[60004] [DocSpec] 하나의 메시지에 신규(OECD1)와 정정(OECD2/OECD3)을 혼합할 수 없습니다."
                );
            }
        }

        private static void ValidateSingleDocSpec(
            Globe.DocSpecType docSpec,
            string section,
            List<string> errors
        )
        {
            if (docSpec == null)
            {
                errors.Add($"[60011] [{section}/DocSpec] DocSpec이 없습니다.");
                return;
            }

            // 60011: DocRefId 형식 [발신관할권국가코드][보고연도][고유식별번호]
            if (string.IsNullOrEmpty(docSpec.DocRefId))
            {
                errors.Add($"[60011] [{section}/DocSpec] DocRefId가 비어 있습니다.");
            }

            // 60012: OECD1/OECD0이면 CorrDocRefId 생략
            if (
                (
                    docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd1
                    || docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd0
                ) && !string.IsNullOrEmpty(docSpec.CorrDocRefId)
            )
            {
                errors.Add(
                    $"[60012] [{section}/DocSpec] DocTypeIndic이 OECD1/OECD0인 경우 CorrDocRefId는 생략되어야 합니다."
                );
            }

            // 60015: OECD2/OECD3이면 CorrDocRefId 필수
            if (
                (
                    docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd2
                    || docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd3
                ) && string.IsNullOrEmpty(docSpec.CorrDocRefId)
            )
            {
                errors.Add(
                    $"[60015] [{section}/DocSpec] DocTypeIndic이 OECD2/OECD3인 경우 CorrDocRefId가 필수입니다."
                );
            }
        }

        #endregion

        #region FilingInfo (60020, 60021, 60022, 60023)

        private static void ValidateFilingInfo(Globe.GlobeOecd globe, List<string> errors)
        {
            var filingInfo = globe.GlobeBody?.FilingInfo;
            if (filingInfo == null)
                return;

            var period = filingInfo.Period;
            if (period != null)
            {
                // 60020: Start ≤ End
                if (period.Start != default && period.End != default && period.Start > period.End)
                {
                    errors.Add(
                        $"[60020] [1.1~1.2 신고기본정보/Period] 기간 시작일({period.Start:yyyy-MM-dd})이 종료일({period.End:yyyy-MM-dd})보다 늦습니다."
                    );
                }

                // 60021: Period End ≤ ReportingPeriod
                if (
                    period.End != default
                    && globe.MessageSpec?.ReportingPeriod != default
                    && period.End > globe.MessageSpec.ReportingPeriod
                )
                {
                    errors.Add(
                        $"[60021] [1.1~1.2 신고기본정보/Period] 기간 종료일({period.End:yyyy-MM-dd})이 ReportingPeriod({globe.MessageSpec.ReportingPeriod:yyyy-MM-dd})보다 늦습니다."
                    );
                }
            }

            // 60023: FilingCE ResCountryCode = TransmittingCountry
            if (
                filingInfo.FilingCe != null
                && globe.MessageSpec != null
                && filingInfo.FilingCe.ResCountryCode != globe.MessageSpec.TransmittingCountry
            )
            {
                errors.Add(
                    $"[60023] [1.1~1.2 신고기본정보/FilingCE] 소재지국({filingInfo.FilingCe.ResCountryCode})이 TransmittingCountry({globe.MessageSpec.TransmittingCountry})와 불일치합니다."
                );
            }

            // 60022: Role=GIR401이면 FilingCE TIN이 UPE TIN 중 하나와 일치해야 함
            if (
                filingInfo.FilingCe?.Role == Globe.FilingCeRoleEnumType.Gir401
                && filingInfo.FilingCe?.Tin != null
            )
            {
                var filingTin = filingInfo.FilingCe.Tin.Value;
                var upeTins = GetUpeTins(globe);

                if (
                    !string.IsNullOrEmpty(filingTin)
                    && upeTins.Count > 0
                    && !upeTins.Contains(filingTin)
                )
                {
                    errors.Add(
                        $"[60022] [1.1~1.2 신고기본정보/FilingCE] Role=GIR401(UPE)인데 FilingCE TIN({filingTin})이 UPE TIN({string.Join(", ", upeTins)}) 중 어느 것과도 일치하지 않습니다."
                    );
                }
            }
        }

        private static List<string> GetUpeTins(Globe.GlobeOecd globe)
        {
            var tins = new List<string>();
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null)
                return tins;

            foreach (var upe in corpStructure.Upe)
            {
                if (upe.OtherUpe?.Id?.Tin != null)
                    tins.AddRange(upe.OtherUpe.Id.Tin.Select(t => t.Value));
                if (upe.ExcludedUpe?.Id?.Tin != null)
                    tins.AddRange(upe.ExcludedUpe.Id.Tin.Select(t => t.Value));
            }

            return tins.Where(t => !string.IsNullOrEmpty(t)).ToList();
        }

        #endregion

        #region RecJurCode (60018)

        private static void ValidateRecJurCode(Globe.GlobeOecd globe, List<string> errors)
        {
            var recJurCodes = globe.GlobeBody?.GeneralSection?.RecJurCode;
            if (recJurCodes == null || recJurCodes.Count == 0)
                return;

            // 60019: Role GIR403/404/405이면 로컬 제출 (교환 불가)
            var role = globe.GlobeBody?.FilingInfo?.FilingCe?.Role;
            if (
                role == Globe.FilingCeRoleEnumType.Gir403
                || role == Globe.FilingCeRoleEnumType.Gir404
                || role == Globe.FilingCeRoleEnumType.Gir405
            )
            {
                errors.Add(
                    $"[60019] [1.1~1.2 신고기본정보/RecJurCode] FilingCE Role({role})은 로컬 제출(Local Lodgement)로, 자동교환 대상이 아닙니다."
                );
            }

            // 60018: RecJurCode 중 하나가 ReceivingCountry와 일치해야 함
            var recvCountry = globe.MessageSpec?.ReceivingCountry;
            if (recvCountry != null && !recJurCodes.Contains(recvCountry.Value))
            {
                errors.Add(
                    $"[60018] [1.1~1.2 신고기본정보/RecJurCode] ReceivingCountry({recvCountry})가 RecJurCode에 포함되어 있지 않습니다."
                );
            }
        }

        #endregion

        #region UPE (70009, 70010)

        private static void ValidateUpe(Globe.GlobeOecd globe, List<string> errors)
        {
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null)
                return;

            // 70009: UPE GlobeStatus에 허용되지 않는 값
            var disallowed = new HashSet<Globe.IdTypeGloBeStatusEnumType>
            {
                Globe.IdTypeGloBeStatusEnumType.Gir305, // PE
                Globe.IdTypeGloBeStatusEnumType.Gir307, // Minority-Owned Parent
                Globe.IdTypeGloBeStatusEnumType.Gir308, // Minority-Owned Subsidiary
                Globe.IdTypeGloBeStatusEnumType.Gir309,
                Globe.IdTypeGloBeStatusEnumType.Gir312,
                Globe.IdTypeGloBeStatusEnumType.Gir313,
                Globe.IdTypeGloBeStatusEnumType.Gir314,
                Globe.IdTypeGloBeStatusEnumType.Gir315,
                Globe.IdTypeGloBeStatusEnumType.Gir317,
                Globe.IdTypeGloBeStatusEnumType.Gir318,
            };

            foreach (var upe in corpStructure.Upe)
            {
                var statuses = new List<Globe.IdTypeGloBeStatusEnumType>();
                if (upe.OtherUpe?.Id?.GlobeStatus != null)
                    statuses.AddRange(upe.OtherUpe.Id.GlobeStatus);
                if (upe.ExcludedUpe?.Id?.GlobeStatus != null)
                    statuses.AddRange(upe.ExcludedUpe.Id.GlobeStatus);

                var upeLabel = upe.OtherUpe?.Id?.Name ?? upe.ExcludedUpe?.Id?.Name ?? "(이름없음)";
                foreach (var status in statuses)
                {
                    if (disallowed.Contains(status))
                    {
                        errors.Add(
                            $"[70009] [1.3.1 최종모기업/'{upeLabel}'] GlobeStatus에 허용되지 않는 값({status})이 포함되어 있습니다."
                        );
                    }
                }

                // 70010: OtherUPE ResCountryCode는 하나만 허용
                if (upe.OtherUpe?.Id?.ResCountryCode?.Count > 1)
                {
                    errors.Add(
                        $"[70010] [1.3.1 최종모기업/'{upeLabel}'] ResCountryCode에 {upe.OtherUpe.Id.ResCountryCode.Count}개 값이 있습니다. 하나만 허용됩니다."
                    );
                }
            }
        }

        #endregion

        #region CE (70011~70032)

        private static void ValidateCe(Globe.GlobeOecd globe, List<string> errors)
        {
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null)
                return;

            foreach (var ce in corpStructure.Ce)
            {
                var ceName = ce.Id?.Name ?? "(이름없음)";
                var loc = $"1.3.2.1 구성기업/'{ceName}'";

                // 70011: CE ResCountryCode 하나만 허용
                if (ce.Id?.ResCountryCode?.Count > 1)
                {
                    errors.Add(
                        $"[70011] [{loc}] ResCountryCode에 {ce.Id.ResCountryCode.Count}개 값이 있습니다. 하나만 허용됩니다."
                    );
                }

                // 70013: GIR313(JV)과 GIR314(JV Subsidiary) 동시 불가
                var statuses = ce.Id?.GlobeStatus;
                if (statuses != null)
                {
                    if (
                        statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir313)
                        && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir314)
                    )
                    {
                        errors.Add(
                            $"[70013] [{loc}] GIR313(JV)과 GIR314(JV Subsidiary)가 동시에 보고되었습니다."
                        );
                    }

                    // 70014: GIR307과 GIR308 동시 불가
                    if (
                        statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir307)
                        && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir308)
                    )
                    {
                        errors.Add(
                            $"[70014] [{loc}] GIR307(소수지분모기업)과 GIR308(소수지분 자회사)가 동시에 보고되었습니다."
                        );
                    }

                    // 70018: GIR305(PE)와 GIR306(Main Entity) 동시 불가
                    if (
                        statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir305)
                        && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir306)
                    )
                    {
                        errors.Add(
                            $"[70018] [{loc}] GIR305(고정사업장)과 GIR306(본점)이 동시에 보고되었습니다."
                        );
                    }

                    // 70020: GIR316/GIR318은 유일한 값이어야 함
                    if (
                        (
                            statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir316)
                            || statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318)
                        )
                        && statuses.Count > 1
                    )
                    {
                        errors.Add(
                            $"[70020] [{loc}] GIR316/GIR318은 GlobeStatus에서 유일한 값이어야 합니다."
                        );
                    }

                    // 70016: GIR307 있으면 GIR309도 필요
                    if (
                        statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir307)
                        && !statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir309)
                    )
                        errors.Add(
                            $"[70016] [{loc}] GIR307(소수지분모기업)이 있으면 GIR309도 함께 보고해야 합니다."
                        );

                    // 70017: GIR308 있으면 GIR309도 필요
                    if (
                        statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir308)
                        && !statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir309)
                    )
                        errors.Add(
                            $"[70017] [{loc}] GIR308(소수지분 자회사)이 있으면 GIR309도 함께 보고해야 합니다."
                        );

                    // 70021: GIR316/318이면 OwnershipChange 필수
                    if (
                        (
                            statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir316)
                            || statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318)
                        )
                        && ce.OwnershipChange.Count == 0
                    )
                        errors.Add(
                            $"[70021] [{loc}] GIR316 또는 GIR318이면 OwnershipChange(1.3.3 시트)가 있어야 합니다."
                        );
                }

                // 70026: GIR305(PE)이면 OwnershipPercentage = 100%
                if (statuses != null && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir305))
                {
                    foreach (var own in ce.Ownership)
                    {
                        if (own.OwnershipPercentage != 1.0m)
                        {
                            errors.Add(
                                $"[70026] [{loc}/Ownership] GIR305(고정사업장)이므로 OwnershipPercentage가 100%여야 합니다. (현재: {own.OwnershipPercentage:P0})"
                            );
                        }
                    }
                }

                // 70027: GIR318이면 OwnershipPercentage=0%, TIN=NOTIN, OwnershipType=GIR806
                if (statuses != null && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318))
                {
                    foreach (var own in ce.Ownership)
                    {
                        if (own.OwnershipPercentage != 0m)
                            errors.Add(
                                $"[70027] [{loc}/Ownership] GIR318이면 OwnershipPercentage=0%이어야 합니다. (현재: {own.OwnershipPercentage:P0})"
                            );
                        if (
                            !string.Equals(
                                own.Tin?.Value,
                                "NOTIN",
                                StringComparison.OrdinalIgnoreCase
                            )
                        )
                            errors.Add(
                                $"[70027] [{loc}/Ownership] GIR318이면 Ownership/TIN='NOTIN'이어야 합니다."
                            );
                        if (own.OwnershipType != Globe.OwnershipTypeEnumType.Gir806)
                            errors.Add(
                                $"[70027] [{loc}/Ownership] GIR318이면 OwnershipType=GIR806이어야 합니다."
                            );
                    }
                }

                // 70028: GIR318이 아니면 OwnershipPercentage는 0% 불가
                if (statuses == null || !statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318))
                {
                    foreach (var own in ce.Ownership)
                    {
                        if (own.OwnershipPercentage == 0m)
                            errors.Add(
                                $"[70028] [{loc}/Ownership] GIR318이 아닌 경우 OwnershipPercentage는 0%가 될 수 없습니다."
                            );
                    }
                }

                // 70032: QIIR 제공 시 Rules에 GIR201/GIR202 필수
                if (ce.Qiir != null && ce.Id?.Rules != null)
                {
                    if (
                        !ce.Id.Rules.Contains(Globe.IdTypeRulesEnumType.Gir201)
                        && !ce.Id.Rules.Contains(Globe.IdTypeRulesEnumType.Gir202)
                    )
                    {
                        errors.Add(
                            $"[70032] [{loc}/QIIR] QIIR가 제공되었으나 Rules에 GIR201/GIR202가 포함되지 않았습니다."
                        );
                    }
                }

                // 70033: QIIR/Exception/TIN은 다른 CE TIN과 일치해야
                if (ce.Qiir?.Exception?.Tin != null)
                {
                    var exTinVal = ce.Qiir.Exception.Tin.Value;
                    if (
                        !string.IsNullOrEmpty(exTinVal)
                        && !string.Equals(exTinVal, "NOTIN", StringComparison.OrdinalIgnoreCase)
                    )
                    {
                        var allCeTins = corpStructure
                            .Ce.Where(c => c != ce)
                            .SelectMany(c => c.Id?.Tin ?? Enumerable.Empty<Globe.TinType>())
                            .Select(t => t.Value)
                            .ToHashSet();
                        if (!allCeTins.Contains(exTinVal))
                            errors.Add(
                                $"[70033] [{loc}/QIIR/Exception] TIN('{exTinVal}')이 기업구조 내 다른 CE의 TIN과 일치하지 않습니다."
                            );
                    }
                }

                // 70034: POPE-IPE=GIR902(IPE)이고 Exception이 있으면 Art2.1.3 선택 필요
                if (
                    ce.Qiir != null
                    && ce.Qiir.PopeIpe == Globe.PopeipeEnumType.Gir902
                    && ce.Qiir.Exception != null
                    && ce.Qiir.Exception.ExceptionRule?.Art213Specified != true
                )
                    errors.Add(
                        $"[70034] [{loc}/QIIR/Exception] POPE-IPE=GIR902(IPE)이고 Exception이 있으면 Art2.1.3 예외를 선택해야 합니다."
                    );
            }

            // 70015: GIR308 존재 시, 다른 CE에 GIR307 필요
            var hasGir308 = corpStructure.Ce.Any(c =>
                c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir308) == true
            );
            var hasGir307 = corpStructure.Ce.Any(c =>
                c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir307) == true
            );
            if (hasGir308 && !hasGir307)
            {
                errors.Add(
                    "[70015] [1.3.2.1 구성기업] GIR308(소수지분 자회사)이 존재하지만 GIR307(소수지분모기업)이 보고된 CE가 없습니다."
                );
            }

            // 70019: GIR305(PE) 존재 시, GIR306(Main Entity)이 다른 CE에 필요
            var hasPe = corpStructure.Ce.Any(c =>
                c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir305) == true
            );
            var hasMain = corpStructure.Ce.Any(c =>
                c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir306) == true
            );
            if (hasPe && !hasMain)
            {
                errors.Add(
                    "[70019] [1.3.2.1 구성기업] GIR305(고정사업장)이 존재하지만 GIR306(본점)을 포함하는 CE가 없습니다."
                );
            }
        }

        #endregion

        #region Rules 일관성 (70012)

        private static void ValidateRulesConsistency(Globe.GlobeOecd globe, List<string> errors)
        {
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null)
                return;

            // 관할지역 → (기업명, Rules 집합) 수집
            var jurMap =
                new Dictionary<
                    string,
                    List<(string Name, HashSet<Globe.IdTypeRulesEnumType> Rules)>
                >();

            void AddEntry(
                Globe.CountryCodeType? country,
                string name,
                IList<Globe.IdTypeRulesEnumType> rules
            )
            {
                if (country == null || rules == null)
                    return;
                var key = country.ToString();
                if (!jurMap.ContainsKey(key))
                    jurMap[key] = new List<(string, HashSet<Globe.IdTypeRulesEnumType>)>();
                jurMap[key].Add((name, new HashSet<Globe.IdTypeRulesEnumType>(rules)));
            }

            // UPE OtherUpe
            foreach (var upe in corpStructure.Upe)
            {
                if (upe.OtherUpe?.Id != null)
                {
                    var country = upe.OtherUpe.Id.ResCountryCode?.FirstOrDefault();
                    AddEntry(country, upe.OtherUpe.Id.Name ?? "(UPE)", upe.OtherUpe.Id.Rules);
                }
                if (upe.ExcludedUpe?.Id != null)
                {
                    var country = upe.ExcludedUpe.Id.ResCountryCode?.FirstOrDefault();
                    AddEntry(
                        country,
                        upe.ExcludedUpe.Id.Name ?? "(ExcludedUPE)",
                        upe.ExcludedUpe.Id.Rules
                    );
                }
            }

            // CE
            foreach (var ce in corpStructure.Ce)
            {
                if (ce.Id == null)
                    continue;
                var country = ce.Id.ResCountryCode?.FirstOrDefault();
                AddEntry(country, ce.Id.Name ?? "(CE)", ce.Id.Rules);
            }

            // 관할지역별 Rules 일관성 검사
            foreach (var (jur, entities) in jurMap)
            {
                if (entities.Count <= 1)
                    continue;

                // GIR204(QDMTT)를 가진 기업이 있으면 해당 관할지역 전체 면제
                bool hasQdmtt = entities.Any(e =>
                    e.Rules.Contains(Globe.IdTypeRulesEnumType.Gir204)
                );
                if (hasQdmtt)
                    continue;

                // Rules 집합이 모두 동일한지 확인 (GIR204 제외 후 비교)
                var first = entities[0].Rules;
                for (int i = 1; i < entities.Count; i++)
                {
                    if (!first.SetEquals(entities[i].Rules))
                    {
                        var names = string.Join(", ", entities.Select(e => e.Name));
                        errors.Add(
                            $"[70012] [1.3.x 기업구조/Rules] 관할지역 {jur} 소속 기업들({names})의 Rules가 서로 다릅니다. 동일 관할지역 기업은 동일한 Rules를 가져야 합니다."
                        );
                        break;
                    }
                }
            }
        }

        #endregion

        #region Summary (70038, 70039, 70040)

        private static void ValidateSummary(Globe.GlobeOecd globe, List<string> errors)
        {
            var body = globe.GlobeBody;
            if (body?.Summary == null || body.Summary.Count == 0)
                return;

            // 보고기간 (Period End 또는 ReportingPeriod)
            var periodEnd = body.FilingInfo?.Period?.End;
            var reportingPeriod = globe.MessageSpec?.ReportingPeriod;
            var effectiveDate = (periodEnd != default ? periodEnd : reportingPeriod) ?? default;

            // UPE 국가 수집 (70040용)
            var upeCountries = new HashSet<Globe.CountryCodeType>();
            var corpStr = body.GeneralSection?.CorporateStructure;
            if (corpStr != null)
            {
                foreach (var upe in corpStr.Upe)
                {
                    if (upe.OtherUpe?.Id?.ResCountryCode != null)
                        foreach (var cc in upe.OtherUpe.Id.ResCountryCode)
                            upeCountries.Add(cc);
                    if (upe.ExcludedUpe?.Id?.ResCountryCode != null)
                        foreach (var cc in upe.ExcludedUpe.Id.ResCountryCode)
                            upeCountries.Add(cc);
                }
            }

            foreach (var summary in body.Summary)
            {
                var jurName =
                    summary.Jurisdiction?.JurisdictionNameSpecified == true
                        ? summary.Jurisdiction.JurisdictionName.ToString()
                        : "(국가 없음)";

                // 70038: GIR1203/1204/1205는 2028-06-30 이후 불가
                if (effectiveDate > new DateTime(2028, 6, 30))
                {
                    foreach (var sh in summary.SafeHarbour)
                    {
                        if (
                            sh == Globe.SafeHarbourEnumType.Gir1203
                            || sh == Globe.SafeHarbourEnumType.Gir1204
                            || sh == Globe.SafeHarbourEnumType.Gir1205
                        )
                            errors.Add(
                                $"[70038] [1.4 요약/{jurName}] SafeHarbour {sh}은 2028.6.30 이후 사용 불가 (전환기 CbCR 적용면제 만료)."
                            );
                    }
                }

                // 70039: GIR1206은 2026-12-31 이후 불가
                if (
                    effectiveDate > new DateTime(2026, 12, 31)
                    && summary.SafeHarbour.Contains(Globe.SafeHarbourEnumType.Gir1206)
                )
                    errors.Add(
                        $"[70039] [1.4 요약/{jurName}] SafeHarbour GIR1206은 2026.12.31 이후 사용 불가 (UTPR 적용면제 만료)."
                    );

                // 70040: GIR1206은 UPE 국가에서만
                if (
                    summary.SafeHarbour.Contains(Globe.SafeHarbourEnumType.Gir1206)
                    && upeCountries.Count > 0
                    && summary.Jurisdiction?.JurisdictionNameSpecified == true
                    && !upeCountries.Contains(summary.Jurisdiction.JurisdictionName)
                )
                    errors.Add(
                        $"[70040] [1.4 요약/{jurName}] GIR1206은 UPE 국가({string.Join("/", upeCountries)})에서만 사용 가능합니다."
                    );
            }
        }

        #endregion

        #region TIN (70001, 70002, 70003, 70005)

        private static void ValidateTin(Globe.GlobeOecd globe, List<string> errors)
        {
            // FilingCE TIN
            ValidateSingleTin(
                globe.GlobeBody?.FilingInfo?.FilingCe?.Tin,
                "1.1~1.2 신고기본정보/FilingCE",
                false,
                errors
            );

            // UPE TINs
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null)
                return;

            foreach (var upe in corpStructure.Upe)
            {
                if (upe.OtherUpe?.Id != null)
                {
                    var upeName = upe.OtherUpe.Id.Name ?? "(이름없음)";
                    if (upe.OtherUpe.Id.Tin.Count == 0)
                        errors.Add(
                            $"[70005] [1.3.1 최종모기업/OtherUPE '{upeName}'] TIN이 없습니다. UPE는 실제 TIN이 필수입니다."
                        );
                    else
                        foreach (var tin in upe.OtherUpe.Id.Tin)
                            ValidateSingleTin(
                                tin,
                                $"1.3.1 최종모기업/OtherUPE '{upeName}'",
                                true,
                                errors
                            );
                }
                if (upe.ExcludedUpe?.Id != null)
                {
                    var upeName = upe.ExcludedUpe.Id.Name ?? "(이름없음)";
                    if (upe.ExcludedUpe.Id.Tin.Count == 0)
                        errors.Add(
                            $"[70005] [1.3.1 최종모기업/ExcludedUPE '{upeName}'] TIN이 없습니다."
                        );
                    else
                        foreach (var tin in upe.ExcludedUpe.Id.Tin)
                            ValidateSingleTin(
                                tin,
                                $"1.3.1 최종모기업/ExcludedUPE '{upeName}'",
                                false,
                                errors
                            );
                }
            }

            // CE TINs
            foreach (var ce in corpStructure.Ce)
            {
                if (ce.Id?.Tin != null)
                {
                    var ceName = ce.Id?.Name ?? "(이름없음)";
                    foreach (var tin in ce.Id.Tin)
                        ValidateSingleTin(tin, $"1.3.2.1 구성기업/CE '{ceName}'", false, errors);

                    // OwnershipChange.PreOwnership.Tin도 검증 (1.3.3 시트 입력)
                    foreach (var change in ce.OwnershipChange)
                        foreach (var preOwn in change.PreOwnership)
                            if (preOwn.Tin != null)
                                ValidateSingleTin(
                                    preOwn.Tin,
                                    $"1.3.3 기업구조변동/CE '{ceName}'/PreOwnership",
                                    false,
                                    errors
                                );

                    // 70005: GIR316/GIR318이 아닌 CE TIN에도 NOTIN 불가
                    bool isNonGroupOrExited =
                        ce.Id?.GlobeStatus?.Any(s =>
                            s == Globe.IdTypeGloBeStatusEnumType.Gir316
                            || s == Globe.IdTypeGloBeStatusEnumType.Gir318
                        ) == true;
                    if (!isNonGroupOrExited)
                    {
                        foreach (var tin in ce.Id.Tin)
                        {
                            if (
                                string.Equals(
                                    tin.Value,
                                    "NOTIN",
                                    StringComparison.OrdinalIgnoreCase
                                )
                                || (
                                    tin.TypeOfTinSpecified
                                    && tin.TypeOfTin == Globe.TinEnumType.Gir3004
                                )
                                || (tin.UnknownSpecified && tin.Unknown)
                            )
                                errors.Add(
                                    $"[70005] [1.3.2.1 구성기업/CE '{ceName}'] GIR316/GIR318 상태가 아닌 CE TIN에는 NOTIN/GIR3004/Unknown=TRUE가 허용되지 않습니다."
                                );
                        }
                    }
                }
            }
        }

        private static void ValidateSingleTin(
            Globe.TinType tin,
            string context,
            bool isUpe,
            List<string> errors
        )
        {
            if (tin == null)
                return;

            var isNoTin = string.Equals(tin.Value, "NOTIN", StringComparison.OrdinalIgnoreCase);
            var isGir3004 = tin.TypeOfTinSpecified && tin.TypeOfTin == Globe.TinEnumType.Gir3004;
            var isUnknown = tin.UnknownSpecified && tin.Unknown;

            // 70001: GIR3004이면 TIN='NOTIN', Unknown=TRUE, IssuedBy 없음
            if (isGir3004 && (!isNoTin || !isUnknown))
            {
                errors.Add(
                    $"[70001] [{context}] TypeOfTIN=GIR3004이면 TIN='NOTIN', Unknown=TRUE이어야 합니다."
                );
            }
            if (isGir3004 && tin.IssuedBySpecified)
            {
                errors.Add(
                    $"[70001] [{context}] TypeOfTIN=GIR3004이면 IssuedBy를 제공하면 안 됩니다."
                );
            }

            // 70002: TIN='NOTIN'이면 GIR3004, Unknown=TRUE, IssuedBy 미제공
            if (isNoTin && (!isGir3004 || !isUnknown || tin.IssuedBySpecified))
            {
                errors.Add(
                    $"[70002] [{context}] TIN='NOTIN'이면 TypeOfTIN=GIR3004, Unknown=TRUE이어야 하고 IssuedBy는 생략해야 합니다."
                );
            }

            // 70003: Unknown=TRUE이면 NOTIN, GIR3004, IssuedBy 미제공
            if (isUnknown && (!isNoTin || !isGir3004 || tin.IssuedBySpecified))
            {
                errors.Add(
                    $"[70003] [{context}] Unknown=TRUE이면 TIN='NOTIN', TypeOfTIN=GIR3004이어야 하고 IssuedBy는 생략해야 합니다."
                );
            }

            // 70005: UPE TIN에는 GIR3004/Unknown=TRUE 불가
            if (isUpe && (isGir3004 || isUnknown))
            {
                errors.Add(
                    $"[70005] [{context}] UPE TIN에는 TypeOfTIN=GIR3004 또는 Unknown=TRUE가 허용되지 않습니다."
                );
            }

            // 70007: TypeOfTIN=GIR3003이면 P2JJYYYYMMDDCCCXXX 형식
            if (
                tin.TypeOfTinSpecified
                && tin.TypeOfTin == Globe.TinEnumType.Gir3003
                && !string.IsNullOrEmpty(tin.Value)
            )
            {
                // P2 + 2자리국가코드 + 8자리날짜 + 3자리그룹코드 + 3자리고유번호 = 18자
                if (!Regex.IsMatch(tin.Value, @"^P2[A-Z]{2}\d{8}[A-Z0-9]{3}[A-Z0-9]{3}$"))
                {
                    errors.Add(
                        $"[70007] [{context}] TypeOfTIN=GIR3003의 TIN 형식이 올바르지 않습니다. 형식: P2[국가코드2자][날짜8자][그룹코드3자][고유번호3자] (현재: '{tin.Value}')"
                    );
                }
            }

            // [형식경고] TIN이 회사명/문장처럼 보이면 경고 (NOTIN 제외)
            // 일반 TIN은 영숫자 + 하이픈/점/슬래시만 사용. 공백·소문자단어는 회사명일 가능성 큼.
            if (!isNoTin && !string.IsNullOrEmpty(tin.Value))
            {
                if (!Regex.IsMatch(tin.Value, @"^[A-Za-z0-9._\-/]+$"))
                {
                    errors.Add(
                        $"[형식경고] [{context}] TIN '{tin.Value}'에 공백/특수문자가 포함됨. 회사명을 입력했는지 확인하세요."
                    );
                }
            }
        }

        #endregion

        #region JurisdictionSection (70044~70098)

        public static void ValidateJurisdictionSection(Globe.GlobeOecd globe, List<string> errors)
        {
            var periodEnd = globe.GlobeBody?.FilingInfo?.Period?.End ?? default;
            var periodStart = globe.GlobeBody?.FilingInfo?.Period?.Start ?? default;

            // 국가별 SafeHarbour 맵 (70045~53용)
            var shMap = new Dictionary<string, HashSet<Globe.SafeHarbourEnumType>>();
            foreach (var s in globe.GlobeBody?.Summary ?? Enumerable.Empty<Globe.SummaryType>())
            {
                if (s.Jurisdiction?.JurisdictionNameSpecified == true)
                {
                    var key = s.Jurisdiction.JurisdictionName.ToString();
                    if (!shMap.ContainsKey(key))
                        shMap[key] = new HashSet<Globe.SafeHarbourEnumType>();
                    foreach (var sh in s.SafeHarbour)
                        shMap[key].Add(sh);
                }
            }

            foreach (var js in globe.GlobeBody.JurisdictionSection)
            {
                var jur = js.Jurisdiction.ToString();
                shMap.TryGetValue(jur, out var jurSh);
                jurSh ??= new HashSet<Globe.SafeHarbourEnumType>();

                foreach (var etr in js.GLoBeTax.Etr)
                {
                    // 70044: ETRStatus는 ETRException 또는 ETRComputation 중 하나 필수
                    if (
                        etr.EtrStatus != null
                        && etr.EtrStatus.EtrException == null
                        && etr.EtrStatus.EtrComputation == null
                    )
                        errors.Add(
                            $"[70044] [{jur}] ETRStatus에 ETRException 또는 ETRComputation 중 하나가 있어야 합니다."
                        );

                    var exc = etr.EtrStatus?.EtrException;
                    var overall = etr.EtrStatus?.EtrComputation?.OverallComputation;

                    // ── SafeHarbour → 필수 요소 검증 ────────────────────────────
                    // 70045: GIR1203/1204/1205 → TransitionalCbCRSafeHarbour 필수
                    if (
                        (
                            jurSh.Contains(Globe.SafeHarbourEnumType.Gir1203)
                            || jurSh.Contains(Globe.SafeHarbourEnumType.Gir1204)
                            || jurSh.Contains(Globe.SafeHarbourEnumType.Gir1205)
                        )
                        && exc?.TransitionalCbCrSafeHarbour == null
                    )
                        errors.Add(
                            $"[70045] [{jur}] SafeHarbour GIR1203/1204/1205가 있으면 ETRException.TransitionalCbCRSafeHarbour 필수입니다."
                        );

                    // 70047: GIR1203 → Revenue 필수
                    if (
                        jurSh.Contains(Globe.SafeHarbourEnumType.Gir1203)
                        && exc?.TransitionalCbCrSafeHarbour != null
                        && string.IsNullOrEmpty(exc.TransitionalCbCrSafeHarbour.Revenue)
                    )
                        errors.Add(
                            $"[70047] [{jur}] SafeHarbour GIR1203이면 TransitionalCbCRSafeHarbour.Revenue 필수입니다."
                        );

                    // 70048: GIR1204 → IncomeTax 필수
                    if (
                        jurSh.Contains(Globe.SafeHarbourEnumType.Gir1204)
                        && exc?.TransitionalCbCrSafeHarbour != null
                        && string.IsNullOrEmpty(exc.TransitionalCbCrSafeHarbour.IncomeTax)
                    )
                        errors.Add(
                            $"[70048] [{jur}] SafeHarbour GIR1204이면 TransitionalCbCRSafeHarbour.IncomeTax 필수입니다."
                        );

                    // 70049: GIR1206 → UTPRSafeHarbour + CITRate 필수
                    if (jurSh.Contains(Globe.SafeHarbourEnumType.Gir1206))
                    {
                        if (exc?.UtprSafeHarbour == null)
                            errors.Add(
                                $"[70049] [{jur}] SafeHarbour GIR1206이면 ETRException.UTPRSafeHarbour 필수입니다."
                            );
                        else if (exc.UtprSafeHarbour.CitRate == 0)
                            errors.Add(
                                $"[70049] [{jur}] SafeHarbour GIR1206이면 UTPRSafeHarbour.CITRate 필수입니다."
                            );
                    }

                    // 70053: GIR1205 → SubstanceExclusion 필수 (NetGlobeIncome > 0인 경우)
                    if (
                        jurSh.Contains(Globe.SafeHarbourEnumType.Gir1205)
                        && overall?.SubstanceExclusion == null
                    )
                    {
                        var ngiTotal = Dec(overall?.NetGlobeIncome?.Total);
                        if (ngiTotal > 0)
                            errors.Add(
                                $"[70053] [{jur}] SafeHarbour GIR1205이면 SubstanceExclusion 필수입니다 (NetGlobeIncome > 0)."
                            );
                    }

                    if (overall == null)
                        continue;

                    // 70054: RevocationYear는 Status=FALSE일 때만 (JurisdictionSection Election)
                    ValidateElectionRevocation(etr.Election, jur, errors);

                    // ── NetGlobeIncome ────────────────────────────────────────
                    // 70060: GIR2025 항목 있으면 IntShippingIncome 필수
                    if (overall.NetGlobeIncome != null)
                    {
                        bool hasGir2025 = overall.NetGlobeIncome.Adjustments.Any(a =>
                            a.AdjustmentItem == Globe.AdjustmentItemEnumType.Gir2025
                        );
                        if (hasGir2025 && overall.NetGlobeIncome.IntShippingIncome == null)
                            errors.Add(
                                $"[70060] [{jur}] NetGlobeIncome에 GIR2025(국제해운소득) 항목이 있으면 IntShippingIncome을 작성해야 합니다."
                            );
                    }

                    // ── AdjustedCoveredTax ────────────────────────────────────
                    if (overall.AdjustedCoveredTax != null)
                    {
                        var act = overall.AdjustedCoveredTax;

                        // 70061: Art4.6.1=TRUE → GIR2711 음수 항목 필수
                        if (etr.Election?.Art461 == true)
                        {
                            bool hasGir2711Neg = act.Adjustments.Any(a =>
                                a.AdjustmentItem == Globe.FinalAdjustedTaxEnumType.Gir2711
                                && Dec(a.Amount) < 0
                            );
                            if (!hasGir2711Neg)
                                errors.Add(
                                    $"[70061] [{jur}] Art4.6.1 선택=TRUE이면 AdjustedCoveredTax에 GIR2711 음수 항목이 있어야 합니다."
                                );
                        }

                        // 70062: GIR2720 있으면 AdjustedCoveredTax 총액 음수 불가
                        bool hasGir2720 = act.Adjustments.Any(a =>
                            a.AdjustmentItem == Globe.FinalAdjustedTaxEnumType.Gir2720
                        );
                        if (hasGir2720 && Dec(act.Total) < 0)
                            errors.Add(
                                $"[70062] [{jur}] AdjustmentItem=GIR2720이 있으면 AdjustedCoveredTax 총액은 음수가 될 수 없습니다."
                            );

                        // ── PostFilingAdjust Year ─────────────────────────────
                        if (act.PostFilingAdjust != null)
                        {
                            var dtAsset = act.PostFilingAdjust.DeferTaxAsset;
                            if (dtAsset != null)
                            {
                                // 70066: Year ≤ Period Start YYYY
                                var dtYears = dtAsset
                                    .AmountAttributed.Where(a => a.Year != default)
                                    .ToList();
                                foreach (var a in dtYears)
                                    if (periodStart != default && a.Year.Year >= periodStart.Year)
                                        errors.Add(
                                            $"[70066] [{jur}] PostFilingAdjust.DeferTaxAsset.AmountAttributed.Year({a.Year:yyyy})은 기간 시작일({periodStart:yyyy}) 이전이어야 합니다."
                                        );
                                // 70067: Year 중복 불가
                                var dupYears = dtYears
                                    .GroupBy(a => a.Year.Year)
                                    .Where(g => g.Count() > 1)
                                    .Select(g => g.Key);
                                foreach (var y in dupYears)
                                    errors.Add(
                                        $"[70067] [{jur}] PostFilingAdjust.DeferTaxAsset.AmountAttributed Year({y})가 중복됩니다."
                                    );
                            }

                            var ctRefund = act.PostFilingAdjust.CoveredTaxRefund;
                            if (ctRefund != null)
                            {
                                // 70068: Year ≤ Period Start YYYY
                                var ctYears = ctRefund
                                    .AmountAttributed.Where(a => a.Year != default)
                                    .ToList();
                                foreach (var a in ctYears)
                                    if (periodStart != default && a.Year.Year >= periodStart.Year)
                                        errors.Add(
                                            $"[70068] [{jur}] PostFilingAdjust.CoveredTaxRefund.AmountAttributed.Year({a.Year:yyyy})은 기간 시작일({periodStart:yyyy}) 이전이어야 합니다."
                                        );
                                // 70069: Year 중복 불가
                                var dupCtYears = ctYears
                                    .GroupBy(a => a.Year.Year)
                                    .Where(g => g.Count() > 1)
                                    .Select(g => g.Key);
                                foreach (var y in dupCtYears)
                                    errors.Add(
                                        $"[70069] [{jur}] PostFilingAdjust.CoveredTaxRefund.AmountAttributed Year({y})가 중복됩니다."
                                    );
                            }
                        }

                        // ── DeemedDistTax Recapture ───────────────────────────
                        if (act.DeemedDistTax?.Election?.Recapture != null)
                        {
                            foreach (var recapture in act.DeemedDistTax.Election.Recapture)
                            {
                                var recYear = recapture.Year;
                                // 70070: Year ≤ Period End
                                if (
                                    periodEnd != default
                                    && recYear != default
                                    && recYear > periodEnd
                                )
                                    errors.Add(
                                        $"[70070] [{jur}] DeemedDistTax.Recapture.Year({recYear:yyyy-MM-dd})은 기간 종료일({periodEnd:yyyy-MM-dd}) 이후일 수 없습니다."
                                    );
                                // 70071: Year ≥ Period End − 3년 (보고FY + 이전 3FY)
                                if (
                                    periodEnd != default
                                    && recYear != default
                                    && recYear.Year < periodEnd.Year - 3
                                )
                                    errors.Add(
                                        $"[70071] [{jur}] DeemedDistTax.Recapture.Year({recYear:yyyy})은 기간 종료일({periodEnd.Year})보다 4년 이상 이전일 수 없습니다."
                                    );

                                // 70072: EndAmount = StartAmount - TotalDDT
                                if (
                                    TryParseDec(recapture.StartAmount, out var start)
                                    && TryParseDec(recapture.TotalDdt, out var total)
                                    && TryParseDec(recapture.EndAmount, out var end)
                                )
                                {
                                    var expected = start - total;
                                    if (Math.Abs(end - expected) > 0.01m)
                                        errors.Add(
                                            $"[70072] [{jur}] DeemedDistTax.Recapture.EndAmount({end})은 StartAmount({start}) - TotalDDT({total}) = {expected}여야 합니다."
                                        );
                                }
                                // 70073: EndAmount 음수 불가
                                if (TryParseDec(recapture.EndAmount, out var endAmt) && endAmt < 0)
                                    errors.Add(
                                        $"[70073] [{jur}] DeemedDistTax.Recapture.EndAmount({endAmt})은 음수일 수 없습니다."
                                    );

                                // 70074: TotalDDT = DDTYear0+1+2+3
                                if (
                                    TryParseDec(recapture.DdtYear0, out var y0)
                                    && TryParseDec(recapture.DdtYear1, out var y1)
                                    && TryParseDec(recapture.DdtYear2, out var y2)
                                    && TryParseDec(recapture.DdtYear3, out var y3)
                                    && TryParseDec(recapture.TotalDdt, out var tDdt)
                                )
                                {
                                    var sumYears = y0 + y1 + y2 + y3;
                                    if (Math.Abs(tDdt - sumYears) > 0.01m)
                                        errors.Add(
                                            $"[70074] [{jur}] DeemedDistTax.Recapture.TotalDDT({tDdt})은 DDTYear 합계({sumYears})와 일치해야 합니다."
                                        );
                                }
                                // 70075: Year=Period End YYYY이면 DDTYear*=0
                                if (periodEnd != default && recYear.Year == periodEnd.Year)
                                {
                                    foreach (
                                        var (label, val) in new[]
                                        {
                                            ("DDTYear-0", recapture.DdtYear0),
                                            ("DDTYear-1", recapture.DdtYear1),
                                            ("DDTYear-2", recapture.DdtYear2),
                                            ("DDTYear-3", recapture.DdtYear3),
                                        }
                                    )
                                        if (Dec(val) != 0)
                                            errors.Add(
                                                $"[70075] [{jur}] DeemedDistTax.Recapture.Year이 기간 종료 연도와 같으면 {label}은 0이어야 합니다."
                                            );
                                }
                            }
                        }

                        // ── TransBlendCFC ─────────────────────────────────────
                        var tblend = act.TransBlendCfc;
                        if (tblend != null && !string.IsNullOrEmpty(tblend.Total))
                        {
                            // 70076: Total = Σ(CfcJur.Allocation.AggAllocTax)
                            var sumAgg = tblend
                                .CfcJur.Where(j => j.Allocation != null)
                                .Sum(j => Dec(j.Allocation.AggAllocTax));
                            if (
                                TryParseDec(tblend.Total, out var tblendTotal)
                                && Math.Abs(tblendTotal - sumAgg) > 0.01m
                            )
                                errors.Add(
                                    $"[70076] [{jur}] TransBlendCFC.Total({tblendTotal})은 CfcJur Allocation.AggAllocTax 합계({sumAgg})와 일치해야 합니다."
                                );
                        }
                    }

                    // ── ExcessNegTaxExpense ───────────────────────────────────
                    var exNeg = overall.ExcessNegTaxExpense;
                    var adjCovTax = overall.AdjustedCoveredTax;
                    if (exNeg != null && adjCovTax != null)
                    {
                        // 70084: GIR2719 Amount = GeneratedInRFY
                        var gir2719 = adjCovTax.Adjustments.FirstOrDefault(a =>
                            a.AdjustmentItem == Globe.FinalAdjustedTaxEnumType.Gir2719
                        );
                        if (
                            gir2719 != null
                            && !string.IsNullOrEmpty(exNeg.GeneratedInRfy)
                            && TryParseDec(gir2719.Amount, out var amt2719)
                            && TryParseDec(exNeg.GeneratedInRfy, out var genRfy)
                            && Math.Abs(amt2719 - genRfy) > 0.01m
                        )
                            errors.Add(
                                $"[70084] [{jur}] AdjustmentItem=GIR2719의 Amount({amt2719})은 ExcessNegTaxExpense.GeneratedInRFY({genRfy})와 일치해야 합니다."
                            );

                        // 70085: GIR2720 Amount = UtilizedInRFY
                        var gir2720 = adjCovTax.Adjustments.FirstOrDefault(a =>
                            a.AdjustmentItem == Globe.FinalAdjustedTaxEnumType.Gir2720
                        );
                        if (
                            gir2720 != null
                            && !string.IsNullOrEmpty(exNeg.UtilizedInRfy)
                            && TryParseDec(gir2720.Amount, out var amt2720)
                            && TryParseDec(exNeg.UtilizedInRfy, out var utilRfy)
                            && Math.Abs(amt2720 - utilRfy) > 0.01m
                        )
                            errors.Add(
                                $"[70085] [{jur}] AdjustmentItem=GIR2720의 Amount({amt2720})은 ExcessNegTaxExpense.UtilizedInRFY({utilRfy})와 일치해야 합니다."
                            );
                    }

                    // ── SubstanceExclusion / ExcessProfits ────────────────────
                    // 70087: SubstanceExclusion.Total = PayrollCost*PayrollMarkUp + TangibleAssetValue*TangibleAssetMarkup
                    if (
                        overall.SubstanceExclusion != null
                        && TryParseDec(overall.SubstanceExclusion.Total, out var sbieTot)
                        && TryParseDec(overall.SubstanceExclusion.PayrollCost, out var payroll)
                        && TryParseDec(
                            overall.SubstanceExclusion.TangibleAssetValue,
                            out var tangible
                        )
                    )
                    {
                        var payMarkup = overall.SubstanceExclusion.PayrollMarkUp;
                        var tangMarkup = overall.SubstanceExclusion.TangibleAssetMarkup;
                        var expectedSbie = payroll * payMarkup + tangible * tangMarkup;
                        if (Math.Abs(sbieTot - expectedSbie) > 1m)
                            errors.Add(
                                $"[70087] [{jur}] SubstanceExclusion.Total({sbieTot})은 PayrollCost×MarkUp + TangibleAsset×MarkUp = {expectedSbie:F2}여야 합니다."
                            );
                    }

                    // 70086: ExcessProfits = max(0, NetGlobeIncome/Total - SubstanceExclusion/Total)
                    if (
                        !string.IsNullOrEmpty(overall.ExcessProfits)
                        && TryParseDec(overall.ExcessProfits, out var exProfit)
                        && TryParseDec(overall.NetGlobeIncome?.Total, out var ngiTot)
                    )
                    {
                        var sbie = Dec(overall.SubstanceExclusion?.Total);
                        var expected86 = Math.Max(0, ngiTot - sbie);
                        if (Math.Abs(exProfit - expected86) > 1m)
                            errors.Add(
                                $"[70086] [{jur}] ExcessProfits({exProfit})은 max(0, NetGlobeIncome.Total({ngiTot}) - SubstanceExclusion.Total({sbie})) = {expected86:F2}여야 합니다."
                            );
                    }

                    // ── Art4.1.5 ─────────────────────────────────────────────
                    if (
                        TryParseDec(overall.NetGlobeIncome?.Total, out var ngiForArt)
                        && ngiForArt < 0
                    )
                    {
                        // 70088: NetGlobeIncome < 0이면 Art4.1.5 필수
                        if (overall.AdditionalTopUpTax?.Art415 == null)
                            errors.Add(
                                $"[70088] [{jur}] NetGlobeIncome.Total({ngiForArt})이 음수이면 AdditionalTopUpTax.Art4.1.5 필수입니다."
                            );
                    }

                    var art415 = overall.AdditionalTopUpTax?.Art415;
                    if (art415 != null)
                    {
                        // 70089: AdjustedCoveredTax 음수여야 함
                        if (
                            TryParseDec(art415.AdjustedCoveredTax, out var art415Act)
                            && art415Act >= 0
                        )
                            errors.Add(
                                $"[70089] [{jur}] Art4.1.5.AdjustedCoveredTax({art415Act})은 음수여야 합니다."
                            );

                        // 70090: GlobeLoss = NetGlobeIncome/Total
                        if (
                            TryParseDec(art415.GlobeLoss, out var globeLoss)
                            && TryParseDec(overall.NetGlobeIncome?.Total, out var ngiCheck)
                            && Math.Abs(globeLoss - ngiCheck) > 0.01m
                        )
                            errors.Add(
                                $"[70090] [{jur}] Art4.1.5.GlobeLoss({globeLoss})은 NetGlobeIncome.Total({ngiCheck})와 일치해야 합니다."
                            );

                        // 70091: ExpectedAdjustedCoveredTax = GlobeLoss × 15%
                        if (
                            TryParseDec(art415.GlobeLoss, out var gl91)
                            && TryParseDec(art415.ExpectedAdjustedCoveredTax, out var eact)
                        )
                        {
                            var exp91 = gl91 * 0.15m;
                            if (Math.Abs(eact - exp91) > 1m)
                                errors.Add(
                                    $"[70091] [{jur}] Art4.1.5.ExpectedAdjustedCoveredTax({eact})은 GlobeLoss({gl91}) × 15% = {exp91:F2}여야 합니다."
                                );
                        }

                        // 70092: AdditionalTopUpTax = max(0, ExpectedAdjustedCoveredTax - AdjustedCoveredTax)
                        if (
                            TryParseDec(art415.ExpectedAdjustedCoveredTax, out var eact92)
                            && TryParseDec(art415.AdjustedCoveredTax, out var act92)
                            && TryParseDec(art415.AdditionalTopUpTax, out var addt92)
                        )
                        {
                            var exp92 = Math.Max(0, eact92 - act92);
                            if (Math.Abs(addt92 - exp92) > 1m)
                                errors.Add(
                                    $"[70092] [{jur}] Art4.1.5.AdditionalTopUpTax({addt92})은 max(0, ExpectedACT({eact92}) - ACT({act92})) = {exp92:F2}여야 합니다."
                                );
                        }
                    }

                    // ── NONArt4.1.5 ──────────────────────────────────────────
                    foreach (
                        var non in overall.AdditionalTopUpTax?.NonArt415
                            ?? Enumerable.Empty<Globe.EtrComputationTypeOverallComputationAdditionalTopUpTaxNonArt415>()
                    )
                    {
                        // 70093: Year ≤ Period End YYYY
                        if (
                            periodEnd != default
                            && non.Year != default
                            && non.Year.Year > periodEnd.Year
                        )
                            errors.Add(
                                $"[70093] [{jur}] NONArt4.1.5.Year({non.Year.Year})은 기간 종료 연도({periodEnd.Year}) 이후일 수 없습니다."
                            );

                        // 70094: Articles=GIR2605 → Year은 최소 4년 이전
                        if (
                            non.Articles.Contains(Globe.NonArt415EnumType.Gir2605)
                            && periodEnd != default
                            && non.Year != default
                            && non.Year.Year > periodEnd.Year - 4
                        )
                            errors.Add(
                                $"[70094] [{jur}] NONArt4.1.5 Articles=GIR2605이면 Year({non.Year.Year})은 기간 종료({periodEnd.Year})보다 최소 4년 이전이어야 합니다."
                            );

                        // 70095: Articles=GIR2602 → Year = 5번째 이전 FY
                        if (
                            non.Articles.Contains(Globe.NonArt415EnumType.Gir2602)
                            && periodEnd != default
                            && non.Year != default
                            && non.Year.Year != periodEnd.Year - 5
                        )
                            errors.Add(
                                $"[70095] [{jur}] NONArt4.1.5 Articles=GIR2602이면 Year({non.Year.Year})은 기간 종료 연도({periodEnd.Year}) - 5 = {periodEnd.Year - 5}이어야 합니다."
                            );

                        // 70096: AdditionalTopUpTax = Recalculated.TopUpTax - Previous.TopUpTax
                        if (
                            non.Previous != null
                            && non.Recalculated != null
                            && TryParseDec(non.Recalculated.TopUpTax, out var recalc)
                            && TryParseDec(non.Previous.TopUpTax, out var prev)
                            && TryParseDec(non.AdditionalTopUpTax, out var addt96)
                        )
                        {
                            var exp96 = recalc - prev;
                            if (Math.Abs(addt96 - exp96) > 1m)
                                errors.Add(
                                    $"[70096] [{jur}] NONArt4.1.5.AdditionalTopUpTax({addt96})은 Recalculated.TopUpTax({recalc}) - Previous.TopUpTax({prev}) = {exp96:F2}여야 합니다."
                                );
                        }
                    }
                }
            }
        }

        // 70054/56: JurisdictionSection Election의 RevocationYear 검증
        private static void ValidateElectionRevocation(
            Globe.EtrTypeElection election,
            string jur,
            List<string> errors
        )
        {
            if (election == null)
                return;

            void Check(string name, bool status, bool rvSpecified)
            {
                if (rvSpecified && status)
                    errors.Add(
                        $"[70054] [{jur}/Election/{name}] RevocationYear는 Status=FALSE(선택 철회)일 때만 제공해야 합니다."
                    );
            }

            if (election.Art321C != null)
                Check(
                    "Art3.2.1.c",
                    election.Art321C.Status,
                    election.Art321C.RevocationYearSpecified
                );
        }

        #endregion

        #region SafeHarbour 완전성 (70041~43)

        private static void ValidateSafeHarbourCompleteness(
            Globe.GlobeOecd globe,
            List<string> errors
        )
        {
            var body = globe.GlobeBody;
            if (body?.Summary == null)
                return;

            var cfs = body.FilingInfo?.AccountingInfo?.CfSofUpe;

            foreach (var summary in body.Summary)
            {
                var jurName =
                    summary.Jurisdiction?.JurisdictionNameSpecified == true
                        ? summary.Jurisdiction.JurisdictionName.ToString()
                        : "(국가 없음)";

                // 70041: CFSofUPE=GIR502/GIR504이면 GIR1207/1208/1209 불가
                if (
                    (
                        cfs == Globe.FilingCeCofUpeEnumType.Gir502
                        || cfs == Globe.FilingCeCofUpeEnumType.Gir504
                    )
                    && summary.SafeHarbour.Any(s =>
                        s == Globe.SafeHarbourEnumType.Gir1207
                        || s == Globe.SafeHarbourEnumType.Gir1208
                        || s == Globe.SafeHarbourEnumType.Gir1209
                    )
                )
                    errors.Add(
                        $"[70041] [1.4 요약/{jurName}] CFSofUPE가 GIR502/GIR504이면 SafeHarbour에 GIR1207/1208/1209를 사용할 수 없습니다."
                    );

                bool hasJurTaxing = summary.JurWithTaxingRightsSpecified;

                // 70042: JurWithTaxingRights 있고 SafeHarbour 없음(또는 GIR1206만) → 4개 필수
                bool onlyGir1206orNone =
                    !summary.SafeHarbour.Any()
                    || (
                        summary.SafeHarbour.Count == 1
                        && summary.SafeHarbour[0] == Globe.SafeHarbourEnumType.Gir1206
                    );
                if (hasJurTaxing && onlyGir1206orNone)
                {
                    if (!summary.EtrRangeSpecified)
                        errors.Add(
                            $"[70042] [1.4 요약/{jurName}] JurWithTaxingRights 작성 시 ETRRange 필수입니다."
                        );
                    if (summary.Sbie == null)
                        errors.Add(
                            $"[70042] [1.4 요약/{jurName}] JurWithTaxingRights 작성 시 SBIE 필수입니다."
                        );
                    if (!summary.QdmtTutSpecified)
                        errors.Add(
                            $"[70042] [1.4 요약/{jurName}] JurWithTaxingRights 작성 시 QDMTTut 필수입니다."
                        );
                    if (!summary.GLoBeTutSpecified)
                        errors.Add(
                            $"[70042] [1.4 요약/{jurName}] JurWithTaxingRights 작성 시 GLoBETut 필수입니다."
                        );
                }

                // 70043: JurWithTaxingRights 있고 SafeHarbour=GIR1202 → ETRRange/SBIE/QDMTTut 필수
                if (hasJurTaxing && summary.SafeHarbour.Contains(Globe.SafeHarbourEnumType.Gir1202))
                {
                    if (!summary.EtrRangeSpecified)
                        errors.Add(
                            $"[70043] [1.4 요약/{jurName}] SafeHarbour=GIR1202이면 ETRRange 필수입니다."
                        );
                    if (summary.Sbie == null)
                        errors.Add(
                            $"[70043] [1.4 요약/{jurName}] SafeHarbour=GIR1202이면 SBIE 필수입니다."
                        );
                    if (!summary.QdmtTutSpecified)
                        errors.Add(
                            $"[70043] [1.4 요약/{jurName}] SafeHarbour=GIR1202이면 QDMTTut 필수입니다."
                        );
                }
            }
        }

        #endregion

        #region Ownership TIN 교차검증 (70030, 70031)

        private static void ValidateCeOwnershipTin(Globe.GlobeOecd globe, List<string> errors)
        {
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null)
                return;

            // 전체 CorporateStructure TIN 목록 (CE + UPE)
            var allTins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var upe in corpStructure.Upe)
            {
                foreach (var t in upe.OtherUpe?.Id?.Tin ?? Enumerable.Empty<Globe.TinType>())
                    if (!string.IsNullOrEmpty(t.Value) && t.Value != "NOTIN")
                        allTins.Add(t.Value);
                foreach (var t in upe.ExcludedUpe?.Id?.Tin ?? Enumerable.Empty<Globe.TinType>())
                    if (!string.IsNullOrEmpty(t.Value) && t.Value != "NOTIN")
                        allTins.Add(t.Value);
            }
            foreach (var ce in corpStructure.Ce)
            foreach (var t in ce.Id?.Tin ?? Enumerable.Empty<Globe.TinType>())
                if (!string.IsNullOrEmpty(t.Value) && t.Value != "NOTIN")
                    allTins.Add(t.Value);

            // GIR306(Main Entity) TIN 목록 (70031용)
            var mainEntityTins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var ce in corpStructure.Ce)
                if (ce.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir306) == true)
                    foreach (var t in ce.Id?.Tin ?? Enumerable.Empty<Globe.TinType>())
                        if (!string.IsNullOrEmpty(t.Value))
                            mainEntityTins.Add(t.Value);

            // OwnershipType: CE(GIR801)/JV(GIR803)/JV Subsidiary(GIR804)는 TIN이 CorporateStructure와 일치해야 함
            var ceOwnerTypes = new HashSet<Globe.OwnershipTypeEnumType>
            {
                Globe.OwnershipTypeEnumType.Gir801,
                Globe.OwnershipTypeEnumType.Gir803,
                Globe.OwnershipTypeEnumType.Gir804,
            };

            foreach (var ce in corpStructure.Ce)
            {
                var ceName = ce.Id?.Name ?? "(이름없음)";
                var loc = $"1.3.2.1 구성기업/'{ceName}'";
                var statuses = ce.Id?.GlobeStatus;

                foreach (var own in ce.Ownership)
                {
                    var tinVal = own.Tin?.Value;
                    if (
                        string.IsNullOrEmpty(tinVal)
                        || tinVal.Equals("NOTIN", StringComparison.OrdinalIgnoreCase)
                    )
                        continue;

                    // 70030: CE/JV/JV Subsidiary OwnershipType이면 TIN이 CorporateStructure에 있어야 함
                    if (ceOwnerTypes.Contains(own.OwnershipType) && !allTins.Contains(tinVal))
                        errors.Add(
                            $"[70030] [{loc}/Ownership] OwnershipType({own.OwnershipType})이 CE/JV인 경우 TIN('{tinVal}')은 기업구조에 신고된 TIN과 일치해야 합니다."
                        );

                    // 70031: GIR305(PE)이면 Ownership TIN이 GIR306(Main Entity) TIN과 일치해야 함
                    if (
                        statuses?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir305) == true
                        && mainEntityTins.Count > 0
                        && !mainEntityTins.Contains(tinVal)
                    )
                        errors.Add(
                            $"[70031] [{loc}/Ownership] GIR305(고정사업장)이면 Ownership TIN('{tinVal}')은 GIR306(본점) CE의 TIN({string.Join("/", mainEntityTins)})과 일치해야 합니다."
                        );
                }
            }
        }

        #endregion

        #region UTPRAttribution (70099~105)

        private static void ValidateUtprAttribution(Globe.GlobeOecd globe, List<string> errors)
        {
            var body = globe.GlobeBody;
            if (body == null)
                return;

            var utprList = body.UtprAttribution;
            if (utprList == null || utprList.Count == 0)
                return;

            var allAttribs = utprList.SelectMany(u => u.Attribution).ToList();

            // 70104: UTPRTopUpTaxCarriedForward 음수 불가
            foreach (var attrib in allAttribs)
            {
                if (TryParseDec(attrib.UtprTopUpTaxCarriedForward, out var carried) && carried < 0)
                    errors.Add(
                        $"[70104] [UTPRAttribution/{attrib.ResCountryCode}] UTPRTopUpTaxCarriedForward({carried})은 음수일 수 없습니다."
                    );

                // 70101: CarryForward ≠ 0이면 Employees 필수
                if (
                    TryParseDec(attrib.UtprTopUpTaxCarryForward, out var carryFwd)
                    && carryFwd != 0
                    && string.IsNullOrEmpty(attrib.Employees)
                )
                    errors.Add(
                        $"[70101] [UTPRAttribution/{attrib.ResCountryCode}] UTPRTopUpTaxCarryForward≠0이면 Employees 필수입니다."
                    );

                // 70102: CarryForward ≠ 0이면 TangibleAssetValue 필수
                if (
                    TryParseDec(attrib.UtprTopUpTaxCarryForward, out var cf2)
                    && cf2 != 0
                    && string.IsNullOrEmpty(attrib.TangibleAssetValue)
                )
                    errors.Add(
                        $"[70102] [UTPRAttribution/{attrib.ResCountryCode}] UTPRTopUpTaxCarryForward≠0이면 TangibleAssetValue 필수입니다."
                    );

                // 70103: CarryForward > 0이면 UTPRPercentage = 0
                if (
                    TryParseDec(attrib.UtprTopUpTaxCarryForward, out var cf3)
                    && cf3 > 0
                    && attrib.UtprPercentage != 0
                )
                    errors.Add(
                        $"[70103] [UTPRAttribution/{attrib.ResCountryCode}] UTPRTopUpTaxCarryForward>0이면 UTPRPercentage는 0%이어야 합니다."
                    );

                // 70105: CarriedForward = CarryForward + Attributed - AddCashTaxExpense
                if (
                    TryParseDec(attrib.UtprTopUpTaxCarryForward, out var cf)
                    && TryParseDec(attrib.UtprTopUpTaxAttributed, out var attr)
                    && TryParseDec(attrib.AddCashTaxExpense, out var cash)
                    && TryParseDec(attrib.UtprTopUpTaxCarriedForward, out var cfwd)
                )
                {
                    var expected = cf + attr - cash;
                    if (Math.Abs(cfwd - expected) > 0.01m)
                        errors.Add(
                            $"[70105] [UTPRAttribution/{attrib.ResCountryCode}] UTPRTopUpTaxCarriedForward({cfwd})은 CarryForward({cf})+Attributed({attr})-AddCashTaxExpense({cash})={expected}여야 합니다."
                        );
                }
            }

            // 70099: Σ(UTPRTopUpTaxAttributed) = Σ(TotalUTPRTopUpTax across all jurisdictions)
            // (TotalUTPRTopUpTax는 JurisdictionSection 내 UTPR 계산 결과 — 현재 미구현 섹션이므로 skip)
        }

        #endregion

        #region 숫자 파싱 헬퍼

        private static bool TryParseDec(string s, out decimal result)
        {
            result = 0;
            if (string.IsNullOrEmpty(s))
                return false;
            return decimal.TryParse(
                s.Replace(",", ""),
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out result
            );
        }

        private static decimal Dec(string s) => TryParseDec(s, out var v) ? v : 0m;

        #endregion
    }
}
