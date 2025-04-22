using Amazon.S3;
using Amazon.S3.Transfer;

using Microsoft.IO;

using SpreadCheetah;

using System;
using System.Threading.Tasks;

class Program
{
    private static readonly RecyclableMemoryStreamManager recyclableMemoryStreamManager = new();

    static async Task Main(string[] args)
    {
        string bucketName = "s3-bucket-local";
        string keyName = "teste_spreadcheetah.xlsx";
        int numberOfSheets = 3;

        var s3Client = new AmazonS3Client("test", "test", new AmazonS3Config
        {
            ServiceURL = "http://localhost:4566",
            ForcePathStyle = true,
        });

        await CreateAndUploadSpreadsheetToS3(s3Client, bucketName, keyName, numberOfSheets);
    }

    static async Task CreateAndUploadSpreadsheetToS3(IAmazonS3 s3Client, string bucketName, string keyName, int numberOfSheets)
    {
        using var memoryStream = recyclableMemoryStreamManager.GetStream();

        var spreadsheet = await Spreadsheet.CreateNewAsync(memoryStream);

        for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++)
        {
            await spreadsheet.StartWorksheetAsync($"Sheet{sheetIndex + 1}");

            for (int rowIndex = 0; rowIndex < 1_000_000; rowIndex++)
            {
                var row = new Cell[40];
                for (int col = 0; col < 40; col++)
                {
                    row[col] = new Cell("test");
                }

                await spreadsheet.AddRowAsync(row);

                if ((rowIndex + 1) % 10_000 == 0)
                {
                    Console.WriteLine($"Sheet {sheetIndex + 1} - linhas escritas: {rowIndex + 1}");
                }
            }
        }

        await spreadsheet.FinishAsync();

        memoryStream.Position = 0;

        var uploadRequest = new TransferUtilityUploadRequest
        {
            InputStream = memoryStream,
            Key = keyName,
            BucketName = bucketName,
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };

        var transferUtility = new TransferUtility(s3Client);
        await transferUtility.UploadAsync(uploadRequest);

        Console.WriteLine("Upload concluído para o S3!");
    }
}
