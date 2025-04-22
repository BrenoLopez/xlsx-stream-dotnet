using SpreadCheetah;

using var fileStream = File.Create("./report.xlsx");

using var spreadsheet = await Spreadsheet.CreateNewAsync(fileStream);

await spreadsheet.StartWorksheetAsync("Sheet 1");

await foreach (var record in GetRecordsAsync())
{
    Cell[] row = [new(record.Description), new(record.Stars)];
    await spreadsheet.AddRowAsync(row);
}

await spreadsheet.FinishAsync();

//using CsvHelper;

//using System.Globalization;

//using var writer = new StreamWriter("./report.csv");

//using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

//await foreach (var record in GetRecordsAsync()) csv.WriteRecord(record);

static async IAsyncEnumerable<SampleRecord> GetRecordsAsync()
{
    await Task.Delay(100);
    yield return new SampleRecord("Lorem ipsum dolor siamet", 5);
}

record SampleRecord(string Description, byte Stars);