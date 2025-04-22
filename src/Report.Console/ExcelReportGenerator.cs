using Microsoft.IO;
using SpreadCheetah;

namespace Report.Console;

public class ExcelReportGenerator
{
    public static async Task<RecyclableMemoryStream> GenerateAsync(IAsyncEnumerable<SampleRecord> records)
    {
        var stream = new RecyclableMemoryStreamManager().GetStream();

        var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        await foreach (var record in records)
        {
            Cell[] row = [new(record.Description), new(record.Stars)];
            await spreadsheet.AddRowAsync(row);
        }

        await spreadsheet.FinishAsync();

        stream.Position = 0;

        return stream;  
    }
}
