//Console.WriteLine("Hello, World!");
using SpreadCheetah;

using var fileStream = File.Create("./report.xlsx");

using var spreadsheet = await Spreadsheet.CreateNewAsync(fileStream);

await spreadsheet.StartWorksheetAsync("Sheet 1");

Cell[] row = [new("Lorem ipsum dolor siamet"), new(10)];

await spreadsheet.AddRowAsync(row);

await spreadsheet.FinishAsync();