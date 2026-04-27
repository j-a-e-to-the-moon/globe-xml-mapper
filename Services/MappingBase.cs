using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public abstract class MappingBase
    {
        protected SheetMapping Mapping { get; }

        protected MappingBase(string mappingFileName)
        {
            if (mappingFileName == null)
                return;
            var jsonPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Resources",
                "mappings",
                mappingFileName
            );
            var json = File.ReadAllText(jsonPath);
            Mapping = JsonSerializer.Deserialize<SheetMapping>(
                json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }
            );
        }

        public string SheetName => Mapping?.SheetName;
        public bool Repeatable => Mapping?.Repeatable ?? false;

        public abstract void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        );

        #region 공통 유틸리티

        protected static void ForEachValue(
            IXLWorksheet ws,
            MappingEntry m,
            string fileName,
            List<string> errors,
            Action<string> action
        )
        {
            try
            {
                var cellValue = ws.Cell(m.Cell).GetString()?.Trim();
                if (string.IsNullOrEmpty(cellValue))
                    return;

                var values = m.Multi
                    ? cellValue.Split(
                        ',',
                        StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries
                    )
                    : new[] { cellValue };

                foreach (var val in values)
                    action(val);
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 셀 {m.Cell} ({m.Label}) 매핑 오류: {ex.Message}");
            }
        }

        protected static void SetEnum<T>(
            string value,
            Action<T> setter,
            List<string> errors,
            string fileName,
            MappingEntry entry
        )
            where T : struct, Enum
        {
            if (TryParseEnum<T>(value, out var result))
                setter(result);
            else
            {
                // 진단: 실제 바이트 값도 포함 (숨은 문자 확인용)
                var bytes = string.Join(
                    " ",
                    System.Text.Encoding.UTF8.GetBytes(value).Select(b => b.ToString("X2"))
                );
                errors.Add(
                    $"[{fileName}] 셀 {entry.Cell}: {typeof(T).Name} 변환 실패 '{value}' (bytes: {bytes})"
                );
            }
        }

        /// <summary>
        /// 날짜 파싱 — 다양한 입력 포맷 흡수.
        ///   1) 표준 날짜 텍스트: "2024-01-01", "2024-01-01 00:00:00"
        ///   2) 4자리 연도: "2024" → 2024-01-01
        ///   3) Excel OADate 시리얼: "45292" → 2024-01-01 (셀 서식이 "일반"인 수식 결과)
        /// </summary>
        protected static bool TryParseDate(string value, out DateTime result)
        {
            result = default;
            if (string.IsNullOrWhiteSpace(value)) return false;
            var trimmed = value.Trim();

            // 표준 날짜 텍스트 우선
            if (DateTime.TryParse(trimmed, out result)) return true;

            // 정수: 1900~9999는 연도로 해석, 그 외(예: 45292)는 Excel OADate 시리얼
            if (int.TryParse(trimmed, out var n))
            {
                if (n >= 1900 && n <= 9999)
                { result = new DateTime(n, 1, 1); return true; }
                if (n >= 1 && n < 100000)
                {
                    try { result = DateTime.FromOADate(n); return true; }
                    catch { /* fall through */ }
                }
            }

            // 소수점 OADate (드물지만 가능)
            if (double.TryParse(trimmed,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var serial)
                && serial >= 1 && serial < 100000)
            {
                try { result = DateTime.FromOADate(serial); return true; }
                catch { /* fall through */ }
            }
            return false;
        }

        /// <summary>
        /// TIN 미보유 시 사용. &lt;TIN TypeOfTIN="GIR3004" unknown="true"&gt;NOTIN&lt;/TIN&gt;
        /// UPE(OtherUPE)에는 사용 불가 (에러 70005) — CE 등 non-UPE 전용.
        /// </summary>
        protected static Globe.TinType NoTin() =>
            new Globe.TinType
            {
                Value = "NOTIN",
                Unknown = true,
                UnknownSpecified = true,
                TypeOfTin = Globe.TinEnumType.Gir3004,
                TypeOfTinSpecified = true,
            };

        /// <summary>
        /// "번호,유형코드,발급국가" 형식 파싱. 유형·발급국가는 생략 가능.
        /// 예: "1234567890,GIR3001,KR"
        /// 빈 값이나 "NOTIN" 입력 시 unknown="true" 처리.
        /// </summary>
        protected static Globe.TinType ParseTin(string input)
        {
            if (
                string.IsNullOrWhiteSpace(input)
                || input.Trim().Equals("NOTIN", StringComparison.OrdinalIgnoreCase)
            )
                return NoTin();

            var parts = input.Split(',', StringSplitOptions.TrimEntries);
            var tin = new Globe.TinType { Value = parts[0] };

            if (
                parts.Length >= 2
                && !string.IsNullOrEmpty(parts[1])
                && TryParseEnum<Globe.TinEnumType>(parts[1], out var tinType)
            )
            {
                tin.TypeOfTin = tinType;
                tin.TypeOfTinSpecified = true;
            }

            if (
                parts.Length >= 3
                && !string.IsNullOrEmpty(parts[2])
                && TryParseEnum<Globe.CountryCodeType>(parts[2], out var country)
            )
            {
                tin.IssuedBy = country;
                tin.IssuedBySpecified = true;
            }

            return tin;
        }

        /// <summary>
        /// "영문;국문" 형식 파싱. ';' 없으면 (value, null).
        /// </summary>
        protected static (string Name, string KName) ParseNameKName(string value)
        {
            var parts = value.Split(';', 2, StringSplitOptions.TrimEntries);
            return parts.Length == 2 ? (parts[0], parts[1]) : (parts[0], null);
        }

        /// <summary>
        /// 빈 문자열/0/공백을 모두 "값 없음"으로 처리.
        /// XSD에서 xs:integer 타입 필드에 빈 값이나 "-0"이 들어가면 검증 실패하므로,
        /// Amount 컬렉션 추가 전 이 메서드로 필터링.
        /// </summary>
        protected static bool HasNonZeroValue(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return false;
            if (decimal.TryParse(raw, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d))
                return d != 0m;
            return true; // 숫자가 아닌 값(드물지만)도 일단 통과
        }

        /// <summary>
        /// 빈 값을 null로 정규화 (XmlSerializer가 ""는 빈 요소로 emit하는 문제 회피).
        /// </summary>
        protected static string NullIfEmpty(string raw) =>
            string.IsNullOrWhiteSpace(raw) ? null : raw;

        /// <summary>
        /// XSD percentage 타입(fractionDigits=4)에 맞춰 소수점 4자리로 반올림.
        /// Excel 수식에서 온 부동소수점 노이즈(0.13598209289311697 등) 제거.
        /// </summary>
        protected static decimal RoundPercentage(decimal value) =>
            System.Math.Round(value, 4, System.MidpointRounding.AwayFromZero);

        /// <summary>
        /// 사용자 입력을 XSD percentage(0~1, fractionDigits=4) decimal로 정규화.
        /// 입력은 반드시 0~1 사이 decimal — 1=100%, 0.5=50%, 0.01=1%, 0=0%.
        /// 범위 밖이거나 파싱 불가/빈 값이면 null 반환 (호출자가 errors.Add 결정).
        /// 예: "0.5" → 0.5 ✓, "1" → 1.0 ✓, "0" → 0 ✓
        ///     "5" → null ❌, "1.5" → null ❌, "5%" → null ❌, "100%" → null ❌
        /// </summary>
        protected static decimal? ParsePercentage(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            if (!decimal.TryParse(raw.Trim(),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var v))
                return null;
            if (v < 0m || v > 1m) return null;
            return RoundPercentage(v);
        }

        /// <summary>
        /// XSD xs:integer 필드에 들어갈 string 값을 정수로 정규화.
        /// "22638.919977616086" → "22639". 빈 값/null/숫자 아님은 그대로 반환.
        /// </summary>
        protected static string RoundToInteger(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return raw;
            if (decimal.TryParse(raw, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d))
                return System.Math.Round(d, 0, System.MidpointRounding.AwayFromZero)
                    .ToString("0", System.Globalization.CultureInfo.InvariantCulture);
            return raw;
        }

        protected static bool ParseBool(string value)
        {
            if (bool.TryParse(value, out var b))
                return b;
            var v = value.Trim().ToUpper();
            return v == "Y"
                || v == "YES"
                || v == "1"
                || v == "TRUE"
                || v == "O"
                || v == "예"
                || v == "여";
        }

        protected static bool TryParseEnum<T>(string value, out T result)
            where T : struct, Enum
        {
            if (TryParseEnumCore<T>(value, out result))
                return true;

            // "GIR701 구성기업" 같이 코드 뒤에 설명이 붙은 경우 첫 단어만 시도
            var firstWord = value.Split(' ', 2)[0].Trim();
            if (
                firstWord.Length > 0
                && firstWord != value
                && TryParseEnumCore<T>(firstWord, out result)
            )
                return true;

            // 비가시 유니코드 문자(NBSP, 제로폭 공백 등) 제거 후 재시도
            var cleaned = new string(
                value
                    .Where(c =>
                        !char.IsControl(c) && c != '\u00A0' && c != '\u200B' && c != '\uFEFF'
                    )
                    .ToArray()
            ).Trim();
            if (cleaned.Length > 0 && cleaned != value && TryParseEnumCore<T>(cleaned, out result))
                return true;

            result = default;
            return false;
        }

        private static bool TryParseEnumCore<T>(string value, out T result)
            where T : struct, Enum
        {
            // XmlEnum 매핑 먼저 시도 — OECD GIR 코드("GIR701" 등)는 XmlEnumAttribute로 정의됨
            // Enum.TryParse의 케이스 인센시티브 동작에 의존하지 않음
            foreach (
                var field in typeof(T).GetFields(
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static
                )
            )
            {
                var xmlAttr = field.GetCustomAttributes(
                    typeof(System.Xml.Serialization.XmlEnumAttribute),
                    false
                );
                if (xmlAttr.Length > 0)
                {
                    var xmlValue = ((System.Xml.Serialization.XmlEnumAttribute)xmlAttr[0]).Name;
                    if (string.Equals(xmlValue, value, StringComparison.OrdinalIgnoreCase))
                    {
                        result = (T)field.GetValue(null);
                        return true;
                    }
                }
            }

            // XmlEnum에 없으면 enum 멤버 이름으로 시도
            // IsDefined 체크 필수 — Enum.TryParse는 "25000" 같은 숫자 문자열도 통과시켜
            // 정의되지 않은 enum 값을 만들고, 직렬화 단계에서 InvalidOperationException 발생
            if (Enum.TryParse(value, true, out result) && Enum.IsDefined(typeof(T), result))
                return true;

            result = default;
            return false;
        }

        #endregion
    }

    #region JSON 모델

    public class SheetMapping
    {
        public string Description { get; set; }
        public string SheetName { get; set; }
        public bool Repeatable { get; set; }
        public string CollectionTarget { get; set; }
        public Dictionary<string, SectionMapping> Sections { get; set; }
    }

    public class SectionMapping
    {
        public string Description { get; set; }
        public List<MappingEntry> Mappings { get; set; }
    }

    public class MappingEntry
    {
        public string Cell { get; set; }
        public string Target { get; set; }
        public string Type { get; set; }
        public string Label { get; set; }
        public bool Multi { get; set; }
    }

    #endregion
}
