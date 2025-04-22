using SpreadCheetah;

class Program
{
    static async Task Main(string[] args)
    {
        string filePath = "./teste_spreadcheetah.xlsx";
        int numberOfSheets = 10;

        await CreateLargeSpreadsheetWithSpreadCheetah(filePath, numberOfSheets);
    }

    static async Task CreateLargeSpreadsheetWithSpreadCheetah(string filePath, int numberOfSheets)
    {
        using (var stream = File.Create(filePath))
        {
            var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

            for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++)
            {
                await spreadsheet.StartWorksheetAsync($"Sheet{sheetIndex + 1}");

                for (int rowIndex = 0; rowIndex < 1_000_000; rowIndex++)
                {
                    var row = new Cell[40];
                    for (int col = 0; col < 40; col++)
                    {
                        row[col] = new("test"); 
                    }

                    await spreadsheet.AddRowAsync(row);

                    if ((rowIndex + 1) % 10_000 == 0)
                    {
                        Console.WriteLine($"Sheet {sheetIndex + 1} - linhas escritas: {rowIndex + 1}");
                    }
                }
            }

            await spreadsheet.FinishAsync();
        }
    }
}
