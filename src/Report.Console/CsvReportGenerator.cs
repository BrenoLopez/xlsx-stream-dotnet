using System.Globalization;

using CsvHelper;

namespace Report.Console;

public class CsvReportGenerator
{
    public static async Task GenerateAsync(IAsyncEnumerable<SampleRecord> records)
    {
        using var writer = new StreamWriter("./report.csv");

        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        await foreach (var record in records) csv.WriteRecord(record);
    }
}
