using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

// 각 Services/Mapping_*.cs 를 분석해 실제 매핑 관계를 JSON 문서로 생성:
//   - 정적 매핑 (cell → target): ws.Cell 호출에서 읽은 값이 XSD 속성으로 흐르는 경로 추적
//   - 동적 블록/헤더 기반 매핑은 "FindRow 헤더" 와 "상대 오프셋" 으로 표기
//
// 출력:
//   Resources/mappings/mapping_{name}.json  — 참조용. 실제 로드 안 됨 (Mapping_1.1~1.2 제외).
//
// 주의: 이 파일은 자동 생성됨. 코드 변경 시 재실행해야 동기화됨.

class Program
{
    static readonly Dictionary<string, (string Sheet, string File, string Section)> MapperInfo = new()
    {
        { "Mapping_1_1_1_2", ("다국적기업그룹 정보", "mapping_1.1~1.2.json", "1.1~1.2") },
        { "Mapping_1_3_1",   ("최종모기업",           "mapping_1.3.1.json",   "1.3.1") },
        { "Mapping_1_3_2_1", ("그룹구조",             "mapping_1.3.2.1.json", "1.3.2.1") },
        { "Mapping_1_3_2_2", ("제외기업",             "mapping_1.3.2.2.json", "1.3.2.2") },
        { "Mapping_1_3_3",   ("그룹구조 변동",        "mapping_1.3.3.json",   "1.3.3") },
        { "Mapping_1_4",     ("요약",                 "mapping_1.4.json",     "1.4") },
        { "Mapping_2",       ("적용면제",             "mapping_2.json",       "2") },
        { "Mapping_JurCal",  ("국가별 계산",          "mapping_jurcal.json",  "3.1~3.3") },
        { "Mapping_EntityCe",("구성기업 계산",        "mapping_entityce.json","3.2.4+3.4") },
        { "Mapping_Utpr",    ("UTPR 배분",            "mapping_utpr.json",    "UTPR") },
    };

    // 파일 → 시트
    static readonly Dictionary<string, string> FileToSheet = new()
    {
        { "Mapping_1.1~1.2.cs",  "다국적기업그룹 정보" },
        { "Mapping_1.3.1.cs",    "최종모기업" },
        { "Mapping_1.3.2.1.cs",  "그룹구조" },
        { "Mapping_1.3.2.2.cs",  "제외기업" },
        { "Mapping_1.3.3.cs",    "그룹구조 변동" },
        { "Mapping_1.4.cs",      "요약" },
        { "Mapping_2.cs",        "적용면제" },
        { "Mapping_JurCal.cs",   "국가별 계산" },
        { "Mapping_EntityCe.cs", "구성기업 계산" },
        { "Mapping_Utpr.cs",     "UTPR 배분" },
    };

    // 시트별 (row, col) → (kind, text)  + (sheet, row) → B열 라벨
    static readonly Dictionary<(string, int, int), (string kind, string text)> CellIndex = new();
    static readonly Dictionary<(string, int), string> RowLabels = new();

    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var templateCsv = Path.Combine(repo, "qa", "template_cells.csv");
        if (File.Exists(templateCsv))
            LoadTemplateCsv(templateCsv);
        else
            Console.WriteLine("[경고] qa/template_cells.csv 없음 — 셀 주소 절대값 계산 제한. DumpTemplate 먼저 실행 권장.");

        var srcDir = Path.Combine(repo, "Services");
        var outDir = Path.Combine(repo, "Resources", "mappings");
        Directory.CreateDirectory(outDir);

        int total = 0;
        foreach (var file in Directory.EnumerateFiles(srcDir, "Mapping_*.cs").OrderBy(p => p))
        {
            var name = Path.GetFileName(file);
            if (!FileToSheet.TryGetValue(name, out var sheet)) continue;
            if (name == "Mapping_1.1~1.2.cs") continue; // 기존 JSON 유지 (실제 로드됨)

            var jsonName = GetJsonName(name);
            var outPath = Path.Combine(outDir, jsonName);
            var count = GenerateFor(file, name, sheet, outPath);
            Console.WriteLine($"  {name} → {jsonName}  ({count}건)");
            total += count;
        }

        Console.WriteLine();
        Console.WriteLine($"총 {total}건 매핑 문서화");
        return 0;
    }

    static string GetJsonName(string csName)
    {
        var stem = Path.GetFileNameWithoutExtension(csName).Replace("Mapping_", "mapping_").ToLowerInvariant();
        return stem + ".json";
    }

    // 한 파일 분석 — method 단위로 블록 잡아 매핑 추출
    static int GenerateFor(string file, string fileName, string sheet, string outPath)
    {
        var text = File.ReadAllText(file);
        var lines = File.ReadAllLines(file);

        var sections = new List<SectionDoc>();
        var methods = ExtractMethods(text, lines);

        foreach (var mth in methods)
        {
            var sec = AnalyzeMethod(mth, lines, sheet);
            if (sec != null && sec.Mappings.Count > 0) sections.Add(sec);
        }

        // FieldMap 배열 기반 매퍼 (1.3.1, 1.3.2.1, 1.3.2.2) 보완
        var fmSection = ExtractFieldMapSection(text, lines, sheet);
        if (fmSection != null && fmSection.Mappings.Count > 0) sections.Add(fmSection);

        // DATA_START_ROW 기반 단순 매퍼 (1.3.3, 1.4, Utpr) 보완
        var dsSection = ExtractDataStartRowSection(text, lines, sheet);
        if (dsSection != null && dsSection.Mappings.Count > 0) sections.Add(dsSection);

        // JSON 직렬화 — mapping_1.1~1.2.json 과 동일한 구조
        var doc = new
        {
            description = sheet,
            sheetName = sheet,
            sourceFile = fileName,
            note = "자동 생성 (GenerateMappingDocs). 개발자 참조용 — 실제 로드 안 됨.",
            sections = DedupSectionKeys(sections).ToDictionary(
                s => s.SectionKey,
                s => new
                {
                    description = s.Description,
                    mappings = s.Mappings.Select(m =>
                    {
                        var d = new Dictionary<string, object?>
                        {
                            ["cell"] = m.Cell,
                            ["target"] = m.Target,
                        };
                        if (!string.IsNullOrEmpty(m.Type)) d["type"] = m.Type;
                        if (!string.IsNullOrEmpty(m.Label)) d["label"] = m.Label;
                        if (m.Multi) d["multi"] = true;
                        return d;
                    }).ToList(),
                }),
        };

        var opts = new JsonSerializerOptions
        {
            WriteIndented = true,
            // '+', '<', '>' 등도 그대로 출력 (JSON 파일은 브라우저 삽입용 아니므로 안전)
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        };
        File.WriteAllText(outPath, JsonSerializer.Serialize(doc, opts), new UTF8Encoding(true));

        return sections.Sum(s => s.Mappings.Count);
    }

    record MethodBlock(string Name, int StartLine, int EndLine);
    record MappingEntry(string Cell, string Target, string Type, string Label, bool Multi);
    class SectionDoc
    {
        public string SectionKey { get; set; } = "";
        public string Description { get; set; } = "";
        public List<MappingEntry> Mappings { get; set; } = new();
    }

    static List<MethodBlock> ExtractMethods(string text, string[] lines)
    {
        var result = new List<MethodBlock>();
        // 메서드 헤더 시작만 감지: `[modifiers] ReturnType MethodName(`
        // 파라미터/닫는 괄호/본문 시작 { 는 멀티라인 허용
        var methodStartRegex = new Regex(
            @"^\s*(?:public|private|protected|internal|static|override|async|\s)+\s*[\w<>?,\s\[\]\.]+?\s+(\w+)\s*\(",
            RegexOptions.Compiled);

        for (int i = 0; i < lines.Length; i++)
        {
            var m = methodStartRegex.Match(lines[i]);
            if (!m.Success) continue;
            var name = m.Groups[1].Value;
            // 예약어/흔한 false positive 제외
            if (name == "if" || name == "for" || name == "foreach" || name == "while" ||
                name == "switch" || name == "return" || name == "using" || name == "throw" ||
                name == "lock" || name == "catch") continue;

            // 이 메서드의 { 본문 시작 찾기 — 이후 라인에서 { 가 처음 나오는 지점
            int bodyStart = -1;
            int parenDepth = 0;
            for (int j = i; j < lines.Length && j < i + 20; j++)
            {
                foreach (var c in lines[j])
                {
                    if (c == '(') parenDepth++;
                    else if (c == ')') parenDepth--;
                    else if (c == '{' && parenDepth == 0) { bodyStart = j; break; }
                }
                if (bodyStart >= 0) break;
            }
            if (bodyStart < 0) continue;

            // 본문 끝 찾기
            int depth = 0;
            int end = -1;
            for (int j = bodyStart; j < lines.Length; j++)
            {
                foreach (var c in lines[j])
                {
                    if (c == '{') depth++;
                    else if (c == '}') { depth--; if (depth == 0) { end = j; break; } }
                }
                if (end >= 0) break;
            }
            if (end > i) { result.Add(new MethodBlock(name, i, end)); i = end; }
        }
        return result;
    }

    // ws.Cell(...) 직후 // 주석 안에서 [A-Z]+[0-9]+ 패턴 (셀주소) 을 모두 찾아 마지막 것을 채택.
    // "// O: +1 국가명 (O5)" → O5 로 인식 (괄호 안의 절대 주소)
    static readonly Regex CellLineRegex = new(
        @"ws\.Cell\s*\(\s*[^,)]+?\s*,\s*\d+\s*\)[^/]*//(.+)$",
        RegexOptions.Compiled);
    static readonly Regex CellAddrRegex = new(
        @"\b([A-Z]+\d+)\b",
        RegexOptions.Compiled);
    static readonly Regex CellAssignRegex = new(
        @"\b(\w+)\s*=\s*ws\.Cell\s*\(\s*([^,)]+?)\s*,\s*(\d+)\s*\)",
        RegexOptions.Compiled);

    static SectionDoc? AnalyzeMethod(MethodBlock mth, string[] lines, string sheet)
    {
        if (mth.Name == "Map" && mth.StartLine < 50) return null;

        var sec = new SectionDoc
        {
            SectionKey = InferSectionKey(mth.Name),
            Description = InferDescription(mth.Name),
        };

        for (int i = mth.StartLine; i <= mth.EndLine && i < lines.Length; i++)
        {
            var line = lines[i];
            var cm = CellLineRegex.Match(line);
            if (!cm.Success) continue;
            // 주석에서 마지막 셀주소 추출 (괄호 안 (O5) 또는 평문 O5)
            var addrMatches = CellAddrRegex.Matches(cm.Groups[1].Value);
            if (addrMatches.Count == 0) continue;
            var cell = addrMatches[addrMatches.Count - 1].Groups[1].Value;
            var am = CellAssignRegex.Match(line);
            var varName = am.Success ? am.Groups[1].Value : "";

            var target = string.IsNullOrEmpty(varName)
                ? TryInferInlineTarget(line)
                : TrackVariableUsage(varName, lines, i, mth.EndLine);

            var type = InferType(varName, lines, i, mth.EndLine);
            var label = GetLabelForCell(sheet, cell);
            var multi = HasSplitUsage(varName, lines, i, mth.EndLine);

            sec.Mappings.Add(new MappingEntry(cell, target, type, label, multi));
        }
        return sec;
    }

    // 셀 참조된 컬럼에 대한 라벨: 같은 행의 B열 라벨 우선, 없으면 인접 라벨 열
    static string GetLabelForCell(string sheet, string cellAddr)
    {
        var m = Regex.Match(cellAddr, @"^([A-Z]+)(\d+)$");
        if (!m.Success) return "";
        var row = int.Parse(m.Groups[2].Value);
        // 같은 행의 라벨 (B열 우선, 그 다음 D, H 등 — 짝이 맞는 라벨 컬럼)
        if (RowLabels.TryGetValue((sheet, row), out var t)) return t;
        // col=2 외의 다른 라벨 컬럼 (D=4, H=8 등)
        for (int c = 2; c <= 20; c++)
        {
            if (CellIndex.TryGetValue((sheet, row, c), out var info) && info.kind == "label")
                return info.text;
        }
        return "";
    }

    // 변수가 Split 으로 여러 값 쪼개지는 지 체크 (multi 플래그)
    static bool HasSplitUsage(string varName, string[] lines, int fromLine, int toLine)
    {
        if (string.IsNullOrEmpty(varName)) return false;
        var re = new Regex(@"\b" + Regex.Escape(varName) + @"\.Split\s*\(", RegexOptions.Compiled);
        for (int i = fromLine + 1; i <= toLine && i < lines.Length; i++)
            if (re.IsMatch(lines[i])) return true;
        return false;
    }

    // 변수 파싱 패턴으로 타입 추정
    static string InferType(string varName, string[] lines, int fromLine, int toLine)
    {
        if (string.IsNullOrEmpty(varName)) return "";
        for (int i = fromLine; i <= toLine && i < lines.Length && i - fromLine < 50; i++)
        {
            var line = lines[i];
            if (!line.Contains(varName)) continue;

            // TryParseEnum<Globe.XEnumType>(varName, ...)
            var tpEnum = Regex.Match(line, @"TryParseEnum\s*<\s*Globe\.(\w+EnumType|CountryCodeType|CurrCodeType)\s*>\s*\(\s*" + Regex.Escape(varName));
            if (tpEnum.Success) return "enum:" + tpEnum.Groups[1].Value;

            // SetEnum<Globe.XEnumType>(varName, ...)
            var setEnum = Regex.Match(line, @"SetEnum\s*<\s*Globe\.(\w+EnumType|CountryCodeType|CurrCodeType)\s*>\s*\(\s*" + Regex.Escape(varName));
            if (setEnum.Success) return "enum:" + setEnum.Groups[1].Value;

            // ParseBool(varName) / TryParseBool
            if (Regex.IsMatch(line, @"\bParseBool\s*\(\s*" + Regex.Escape(varName))) return "bool";

            // TryParseDate(varName, ...)
            if (Regex.IsMatch(line, @"\bTryParseDate\s*\(\s*" + Regex.Escape(varName))) return "date";

            // decimal.TryParse(varName, ...) or int.TryParse
            if (Regex.IsMatch(line, @"\b(?:decimal|int|double|float)\.TryParse\s*\(\s*" + Regex.Escape(varName))) return "decimal";

            // ParseTin(varName) → TIN type
            if (Regex.IsMatch(line, @"\bParseTin\s*\(\s*" + Regex.Escape(varName))) return "tin";
        }
        return "string"; // 기본 — GetString 결과
    }

    // 변수 사용 추적 — 재귀적으로 "이 변수가 최종 XSD 속성에 할당되는 지점" 찾기
    // 깊이 제한: 3 (x → y → z → prop)
    static string TrackVariableUsage(string varName, string[] lines, int fromLine, int toLine, int depth = 0)
    {
        if (depth > 3) return "(깊이초과)";
        var results = new List<string>();

        for (int i = fromLine + 1; i <= toLine && i < lines.Length; i++)
        {
            var line = lines[i];
            if (!Regex.IsMatch(line, @"\b" + Regex.Escape(varName) + @"\b")) continue;
            if (i - fromLine > 40) break; // 너무 먼 라인은 스킵

            // pattern 1: foo.bar.baz = ... varName ...  (직접/변환 포함)
            var assign = Regex.Match(line, @"^\s*([a-zA-Z_]\w*(?:\.\w+)+)\s*(?:\?\?)?=(?!=).*?\b" + Regex.Escape(varName) + @"\b");
            if (assign.Success) { results.Add(assign.Groups[1].Value); break; }

            // pattern 2: v => foo.bar = v   (SetEnum 같은 콜백)
            var cb = Regex.Match(line, @"v\s*=>\s*([a-zA-Z_]\w*(?:\.\w+)+)\s*=\s*v");
            if (cb.Success && line.Contains(varName)) { results.Add(cb.Groups[1].Value); break; }

            // pattern 3: foo.Add(varName) / foo.Add(ParseXxx(varName))
            var addCall = Regex.Match(line, @"([a-zA-Z_]\w*(?:\.\w+)+)\.Add\s*\(");
            if (addCall.Success && line.Contains(varName)) { results.Add(addCall.Groups[1].Value + "[]"); break; }

            // pattern 4: 변환 — var next = TryParse(varName, ...) / ParseTin(varName) / ...
            //   next 변수 이름 찾고 재귀적으로 추적
            var transform = Regex.Match(line,
                @"(?:var|int|decimal|string|bool)\s+(\w+)\s*=\s*(?:Parse\w+|TryParse\w*|[A-Z]\w*)\s*\([^)]*\b" + Regex.Escape(varName) + @"\b");
            if (transform.Success)
            {
                var next = transform.Groups[1].Value;
                var deeper = TrackVariableUsage(next, lines, i, toLine, depth + 1);
                if (deeper != "(unresolved)" && deeper != "(깊이초과)") { results.Add(deeper); break; }
            }

            // pattern 5: TryParseEnum<T>(varName, out var next) — out 변수 추적
            var tryParse = Regex.Match(line,
                @"TryParse\w*\s*(?:<[^>]+>)?\s*\(\s*" + Regex.Escape(varName) + @"\b[^,]*,\s*out\s+(?:var\s+)?(\w+)");
            if (tryParse.Success)
            {
                var next = tryParse.Groups[1].Value;
                var deeper = TrackVariableUsage(next, lines, i, toLine, depth + 1);
                if (deeper != "(unresolved)" && deeper != "(깊이초과)") { results.Add(deeper); break; }
            }

            // pattern 6: 초기화자 { Prop = varName } — 생성하는 타입 컨텍스트 확인
            var init = Regex.Match(line, @"(\w+)\s*=\s*[^;,]*\b" + Regex.Escape(varName) + @"\b");
            if (init.Success && char.IsUpper(init.Groups[1].Value[0]))
            {
                // 근처 new TypeName 찾기 (3라인 위까지)
                string typeName = "<init>";
                for (int k = i; k >= Math.Max(fromLine, i - 3); k--)
                {
                    var tm = Regex.Match(lines[k], @"new\s+(?:Globe\.)?([A-Z]\w*)");
                    if (tm.Success) { typeName = tm.Groups[1].Value; break; }
                }
                results.Add($"{typeName}.{init.Groups[1].Value}");
                break;
            }
        }
        return results.Count > 0 ? results[0] : "(unresolved)";
    }

    // 추적 실패 시: 변수가 처음 사용되는 라인의 간략한 코드 텍스트 반환 (메서드 끝까지)
    static string FindVarUsageSnippet(string varName, string[] lines, int fromLine, int toLine)
    {
        for (int i = fromLine + 1; i <= toLine && i < lines.Length; i++)
        {
            if (!Regex.IsMatch(lines[i], @"\b" + Regex.Escape(varName) + @"\b")) continue;
            var t = lines[i].Trim();
            if (t.Length > 100) t = t.Substring(0, 100) + "…";
            return $"L{i + 1}: {t}";
        }
        return "";
    }

    // 한 라인 안에 LHS 가 바로 있는 경우 (ws.Cell 결과를 변수로 받지 않고 직접 전달)
    static string TryInferInlineTarget(string line)
    {
        var m = Regex.Match(line, @"^\s*([a-zA-Z_]\w*(?:\.\w+)+)\s*=\s*ws\.Cell");
        if (m.Success) return m.Groups[1].Value;
        var add = Regex.Match(line, @"([a-zA-Z_]\w*(?:\.\w+)+)\.Add\s*\(\s*ws\.Cell");
        if (add.Success) return add.Groups[1].Value + "[]";
        return "(inline)";
    }

    static string InferSectionKey(string methodName)
    {
        // Map321OverallComputation → "3.2.1"
        // Map322Transition → "3.2.2.3"
        // 그대로 매핑하긴 어려움 — 그냥 메서드명 사용
        var m = Regex.Match(methodName, @"Map(\d+)([A-Z]\w*)?");
        if (m.Success && m.Groups[1].Value.Length >= 2)
        {
            var digits = m.Groups[1].Value;
            if (digits.Length == 2) return $"{digits[0]}.{digits[1]}";
            if (digits.Length == 3) return $"{digits[0]}.{digits[1]}.{digits[2]}";
            if (digits.Length == 4) return $"{digits[0]}.{digits[1]}.{digits[2]}.{digits[3]}";
        }
        return methodName;
    }

    static List<SectionDoc> DedupSectionKeys(List<SectionDoc> sections)
    {
        var seen = new Dictionary<string, int>();
        var result = new List<SectionDoc>();
        foreach (var s in sections)
        {
            var key = s.SectionKey;
            if (seen.TryGetValue(key, out var n))
            {
                seen[key] = n + 1;
                s.SectionKey = $"{key} ({n + 1})";
            }
            else seen[key] = 1;
            result.Add(s);
        }
        return result;
    }

    static void LoadTemplateCsv(string path)
    {
        var lines = File.ReadAllLines(path);
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = ParseCsvLine(lines[i]);
            if (parts.Count < 8) continue;
            var sheet = parts[0];
            if (!int.TryParse(parts[2], out var row)) continue;
            if (!int.TryParse(parts[3], out var col)) continue;
            CellIndex[(sheet, row, col)] = (parts[6], parts[7]);
            if (col == 2 && parts[6] == "label" && !RowLabels.ContainsKey((sheet, row)))
                RowLabels[(sheet, row)] = parts[7];
        }
    }

    static List<string> ParseCsvLine(string line)
    {
        var result = new List<string>();
        var sb = new StringBuilder();
        bool inQuotes = false;
        for (int i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (inQuotes)
            {
                if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                else if (c == '"') inQuotes = false;
                else sb.Append(c);
            }
            else
            {
                if (c == ',') { result.Add(sb.ToString()); sb.Clear(); }
                else if (c == '"') inQuotes = true;
                else sb.Append(c);
            }
        }
        result.Add(sb.ToString());
        return result;
    }

    static string ColLetter(int col)
    {
        var sb = new StringBuilder();
        while (col > 0) { col--; sb.Insert(0, (char)('A' + col % 26)); col /= 26; }
        return sb.ToString();
    }

    // ─── FieldMap (Mapping_1.3.1 / 1.3.2.1 / 1.3.2.2 같은 패턴) ───────────
    //   private static readonly (int Offset, string Target)[] FieldMap = { (N, "X.Y.Z"), ... };
    //   foreach (var (offset, target) in FieldMap)
    //       ws.Cell(blockStartRow + offset, COL).GetString();
    //   여기서 BLOCK_HEADER 또는 BLOCK_START 로 행 기준 결정.
    static SectionDoc? ExtractFieldMapSection(string text, string[] lines, string sheet)
    {
        var fmRegex = new Regex(
            @"FieldMap\s*=\s*\{([^}]+)\};",
            RegexOptions.Compiled | RegexOptions.Singleline);
        var fm = fmRegex.Match(text);
        if (!fm.Success) return null;

        var body = fm.Groups[1].Value;
        var bodyStartIdx = fm.Groups[1].Index;

        // 엔트리 추출 — (offset, "target") 패턴, 위치도 함께 기록
        var entryRegex = new Regex(@"\(\s*(\d+)\s*,\s*""([^""]+)""\s*\)", RegexOptions.Compiled);
        var entries = entryRegex.Matches(body)
            .Select(m => (
                Offset: int.Parse(m.Groups[1].Value),
                Target: m.Groups[2].Value,
                Line: 1 + text.Substring(0, bodyStartIdx + m.Index).Count(c => c == '\n')
            ))
            .ToList();
        if (entries.Count == 0) return null;

        // BLOCK_HEADER 또는 BLOCK_START 로 첫 블록 시작 행 결정
        int baseRow = 0;
        string baseDesc = "";
        var bh = Regex.Match(text, @"BLOCK_HEADER\s*=\s*""([^""]+)""");
        if (bh.Success)
        {
            var header = bh.Groups[1].Value;
            // sheet 컨텍스트 필요 — caller 에서 전달받아야 하지만, 여기서는 글로벌 시트 맵 접근
            // 현재 파일의 시트는 outer context. 간단화를 위해 모든 시트에서 검색
            foreach (var kv in RowLabels)
                if (kv.Key.Item1 == sheet && kv.Value.Contains(header)) { baseRow = kv.Key.Item2; break; }
            baseDesc = $"BLOCK_HEADER=\"{header}\" (템플릿 행 {baseRow})";
        }
        var bsConst = Regex.Match(text, @"BLOCK_START\s*=\s*(\d+)");
        if (baseRow == 0 && bsConst.Success)
        {
            baseRow = int.Parse(bsConst.Groups[1].Value);
            baseDesc = $"BLOCK_START={baseRow}";
        }

        // col 감지: FieldMap 사용 foreach 블록 안에서 ws.Cell 호출 찾기
        int col = 0;
        var foreachFm = Regex.Match(text, @"foreach\s*\(\s*var\s+\([^)]*\)\s+in\s+FieldMap\s*\)");
        if (foreachFm.Success)
        {
            var searchFrom = foreachFm.Index;
            var searchLen = System.Math.Min(2000, text.Length - searchFrom);
            var window = text.Substring(searchFrom, searchLen);
            var cellM = Regex.Match(window, @"ws\.Cell\s*\(\s*[^,)]+?\s*,\s*(\d+)\s*\)");
            if (cellM.Success) int.TryParse(cellM.Groups[1].Value, out col);
        }
        if (col == 0) col = 15;

        // 블록 크기 / 간격 추정 — SET_SIZE 또는 BLOCK_END-BLOCK_START+GAP
        int blockStride = 0;
        var setSizeM = Regex.Match(text, @"SET_SIZE\s*=\s*(\d+)");
        if (setSizeM.Success) blockStride = int.Parse(setSizeM.Groups[1].Value);

        var sec = new SectionDoc
        {
            SectionKey = bh.Success ? bh.Groups[1].Value : "FieldMap",
            Description = $"블록 반복 ({baseDesc}, 입력 열 = {ColLetter(col)}" +
                          (blockStride > 0 ? $", 블록당 {blockStride}행" : "") + ")",
        };

        foreach (var (offset, target, line) in entries)
        {
            var absCell = baseRow > 0 ? $"{ColLetter(col)}{baseRow + offset}" : $"{ColLetter(col)}?";
            var label = GetLabelForCell(sheet, absCell);
            var type = InferTypeFromTarget(target);
            sec.Mappings.Add(new MappingEntry(absCell, target, type, label, false));
        }

        // 추가: FieldMap 외부의 별도 offset 상수 감지 (예: OWNERSHIP_ROW_OFFSET)
        var extraOffsetRegex = new Regex(
            @"private\s+const\s+int\s+(\w+_OFFSET)\s*=\s*(\d+)\s*;",
            RegexOptions.Compiled);
        foreach (Match em in extraOffsetRegex.Matches(text))
        {
            var constName = em.Groups[1].Value;
            if (constName == "SET_SIZE" || constName == "SET_GAP" ||
                constName.EndsWith("_SIZE") || constName.EndsWith("_GAP") ||
                constName.StartsWith("BLOCK")) continue;
            var offset = int.Parse(em.Groups[2].Value);
            if (baseRow <= 0) continue;
            var absCell = $"{ColLetter(col)}{baseRow + offset}";
            var label = GetLabelForCell(sheet, absCell);
            sec.Mappings.Add(new MappingEntry(absCell,
                $"(별도 처리: {constName})", "complex", label, true));
        }
        return sec;
    }

    // 타겟 이름에서 대략 타입 추정 (FieldMap 용: 파싱 코드 없음)
    static string InferTypeFromTarget(string target)
    {
        if (target.EndsWith(".ResCountryCode")) return "enum:CountryCodeType";
        if (target.EndsWith(".Role")) return "enum:FilingCeRoleEnumType";
        if (target.EndsWith(".Rules")) return "enum:IdTypeRulesEnumType";
        if (target.EndsWith(".GlobeStatus")) return "enum:IdTypeGloBeStatusEnumType";
        if (target.EndsWith(".ExcludedUpeStatus")) return "enum:ExcludedUpeEnumType";
        if (target.EndsWith(".ChangeFlag") || target.EndsWith(".Change") ||
            target.EndsWith(".Art1035") || target.EndsWith("UnreportChangeCorpStr") ||
            target.EndsWith(".Art93") || target.EndsWith(".UpeOwnership")) return "bool";
        if (target.EndsWith(".Tin.Value") || target.EndsWith(".ReceivingTin")) return "string";
        if (target.Contains("Tin")) return "tin";
        return "string";
    }

    static SectionDoc? ExtractDataStartRowSection(string text, string[] lines, string sheet)
    {
        var ds = Regex.Match(text, @"DATA_START_ROW\s*=\s*(\d+)");
        if (!ds.Success) return null;
        var startRow = int.Parse(ds.Groups[1].Value);

        var entries = new List<(int Col, string LhsHint, string Target, int Line)>();
        var rowCellRegex = new Regex(@"ws\.Cell\s*\(\s*row\s*,\s*(\d+)\s*\)", RegexOptions.Compiled);
        for (int i = 0; i < lines.Length; i++)
        {
            foreach (Match m in rowCellRegex.Matches(lines[i]))
            {
                var col = int.Parse(m.Groups[1].Value);
                var assignM = Regex.Match(lines[i], @"^\s*var\s+(\w+)\s*=\s*ws\.Cell");
                var hint = assignM.Success ? assignM.Groups[1].Value : "(direct)";
                // 직접 할당 — foo.bar = ws.Cell(...) 패턴
                var directM = Regex.Match(lines[i], @"^\s*([a-zA-Z_]\w*(?:\.\w+)+)\s*=\s*ws\.Cell");
                var target = directM.Success ? directM.Groups[1].Value : "(unresolved)";
                // var 경우 다음 라인에서 사용 추적
                if (!directM.Success && assignM.Success)
                    target = TrackVariableUsage(hint, lines, i, System.Math.Min(lines.Length - 1, i + 25));
                entries.Add((col, hint, target, i + 1));
            }
        }
        if (entries.Count == 0) return null;

        var sec = new SectionDoc
        {
            SectionKey = "rows",
            Description = $"행 반복: 행 {startRow}부터 EOF까지 각 행에서 여러 열 읽음",
        };
        foreach (var (col, hint, target, line) in entries)
        {
            var cell = $"{ColLetter(col)}{startRow}";
            var label = GetLabelForCell(sheet, cell);
            var type = InferTypeFromTarget(target);
            sec.Mappings.Add(new MappingEntry(cell, target, type, label, false));
        }
        return sec;
    }

    static string InferDescription(string methodName)
    {
        // CamelCase → 분리 (기본 힌트)
        var parts = Regex.Matches(methodName, @"[A-Z][a-z0-9]*")
                        .Select(m => m.Value).ToArray();
        return string.Join(" ", parts);
    }
}
