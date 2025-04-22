namespace Report.Console;

public class ReportInMemoryRepository
{
    public static async IAsyncEnumerable<SampleRecord> GetRecordsAsync(int totalItems)
    {
        await Task.Delay(10);

        for (int i = 0; i < totalItems; i++)
        {
            yield return new SampleRecord($"Record {i + 1}", (byte)(i % 5));
        }
    }
}
