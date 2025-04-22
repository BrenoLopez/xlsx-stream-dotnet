using Amazon.S3;
using Amazon.S3.Model;

using Microsoft.IO;

using SpreadCheetah;

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

class Program
{
    private static readonly RecyclableMemoryStreamManager recyclableMemoryStreamManager = new();

    static async Task Main(string[] args)
    {
        string bucketName = "s3-bucket-local";
        string keyName = "pasta/teste_spreadcheetah_multipart.xlsx";
        int numberOfSheets = 3; 

        var s3Client = new AmazonS3Client("test", "test", new AmazonS3Config
        {
            ServiceURL = "http://localhost:4566",
            ForcePathStyle = true,
        });

        await CreateAndUploadSpreadsheetMultipartAsync(s3Client, bucketName, keyName, numberOfSheets);
    }

    static async Task CreateAndUploadSpreadsheetMultipartAsync(IAmazonS3 s3Client, string bucketName, string keyName, int numberOfSheets)
    {
        using var memoryStream = recyclableMemoryStreamManager.GetStream();

        var options = new SpreadCheetahOptions();
        var spreadsheet = await Spreadsheet.CreateNewAsync(memoryStream, options);

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

        const int partSize = 5 * 1024 * 1024;
        var initRequest = new InitiateMultipartUploadRequest
        {
            BucketName = bucketName,
            Key = keyName,
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };

        var initResponse = await s3Client.InitiateMultipartUploadAsync(initRequest);
        var uploadId = initResponse.UploadId;
        var partETags = new List<PartETag>();
        int partNumber = 1;

        try
        {
            byte[] buffer = new byte[partSize];
            int bytesRead;
            while ((bytesRead = await memoryStream.ReadAsync(buffer, 0, partSize)) > 0)
            {
                using var partStream = new MemoryStream(buffer, 0, bytesRead);

                var uploadRequest = new UploadPartRequest
                {
                    BucketName = bucketName,
                    Key = keyName,
                    UploadId = uploadId,
                    PartNumber = partNumber,
                    InputStream = partStream,
                    IsLastPart = false,
                };

                var uploadResponse = await s3Client.UploadPartAsync(uploadRequest);

                partETags.Add(new PartETag(partNumber, uploadResponse.ETag));
                Console.WriteLine($"Parte {partNumber} enviada, ETag: {uploadResponse.ETag}");

                partNumber++;
            }

            var completeRequest = new CompleteMultipartUploadRequest
            {
                BucketName = bucketName,
                Key = keyName,
                UploadId = uploadId,
                PartETags = partETags
            };

            await s3Client.CompleteMultipartUploadAsync(completeRequest);

            Console.WriteLine("Upload multipart completo!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro no upload: {ex.Message}");
            await s3Client.AbortMultipartUploadAsync(new AbortMultipartUploadRequest
            {
                BucketName = bucketName,
                Key = keyName,
                UploadId = uploadId
            });

            throw;
        }
    }
}
