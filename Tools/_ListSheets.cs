using ClosedXML.Excel;
var path = args[0];
var wb = new XLWorkbook(path);
foreach (var ws in wb.Worksheets)
    Console.WriteLine($"  {ws.Position}: '{ws.Name}' ({ws.Visibility})");
