using Report.Console;

var records = GetRecordsAsync();

await ExcelReportGenerator.GenerateAsync(records);

static async IAsyncEnumerable<SampleRecord> GetRecordsAsync()
{
    await Task.Delay(100);
    yield return new SampleRecord("Lorem ipsum dolor siamet", 5);
}
