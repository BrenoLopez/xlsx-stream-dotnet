namespace Report.Console.Contracts.Input;

public class ReportInMemoryRepository
{
    public static async IAsyncEnumerable<InputRecord> GetRecordsAsync(int totalItems)
    {
        await Task.Delay(10);

        for (int i = 0; i < totalItems; i++)
        {
            yield return new InputRecord(
                Description: $"Record {i + 1}",
                Stars: (byte)(i % 5),
                Title: $"Title {i + 1}",
                Category: $"Category {(i % 3) + 1}",
                CreatedAt: DateTime.UtcNow,
                UpdatedAt: DateTime.UtcNow,
                Author: $"Author {i + 1}",
                IsActive: i % 2 == 0,
                Priority: i % 10,
                Budget: 1000 + i * 10,
                Identifier: Guid.NewGuid(),
                Tags: $"tag{i % 5}",
                Comments: $"This is comment {i + 1}",
                Rating: (i % 5) + 0.5,
                Views: i * 10
            );
        }
    }
}
