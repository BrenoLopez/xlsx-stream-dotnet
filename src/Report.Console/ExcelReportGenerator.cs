using SpreadCheetah;

namespace Report.Console;

public class ExcelReportGenerator
{
    public static async Task GenerateAsync(IAsyncEnumerable<SampleRecord> records)
    {
        using var fileStream = File.Create("./report.xlsx");

        using var spreadsheet = await Spreadsheet.CreateNewAsync(fileStream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        await foreach (var record in records)
        {
            Cell[] row = [new(record.Description), new(record.Stars)];
            await spreadsheet.AddRowAsync(row);
        }

        await spreadsheet.FinishAsync();
    }
}
