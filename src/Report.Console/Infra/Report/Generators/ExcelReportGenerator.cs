using System.Buffers;

using DocumentFormat.OpenXml.Drawing.Charts;

using Microsoft.IO;

using Report.Console.Contracts.Input;

using SpreadCheetah;

using static Microsoft.IO.RecyclableMemoryStreamManager;

namespace Report.Console.Infra.Report.Generators;

public class ExcelReportGenerator(IAmazonS3 s3Client)
{
    const string TemporaryDirectoryPath = "./tmp";
    const int DefaultBufferSize = 81_920;
    const int AllowedItemsPerSheet = 1_000_000;

    public async Task GenerateAsync(string fileName, IAsyncEnumerable<InputRecord> records, CancellationToken cancellationToken = default)
    {

        string filePath = $"{TemporaryDirectoryPath}/{fileName}";

        try
        {
            Directory.CreateDirectory(TemporaryDirectoryPath);

            await using FileStream stream = new(filePath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: DefaultBufferSize, useAsync: true);


            await using Spreadsheet spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: cancellationToken);
            int currentSheet = 1;
            int totalItems = 0;

            await spreadsheet.StartWorksheetAsync($"Sheet {currentSheet}", token: cancellationToken);

            var headers = typeof(InputRecord).GetProperties()
                .Select(prop => new Cell(prop.Name))
                .ToArray();

            await spreadsheet.AddRowAsync(headers, cancellationToken);

            totalItems++;

            await foreach (var record in records.WithCancellation(cancellationToken))
            {
                if (totalItems > 0 && totalItems % AllowedItemsPerSheet == 0)
                {
                    currentSheet++;
                    await spreadsheet.StartWorksheetAsync($"Sheet {currentSheet}", token: cancellationToken);
                }


                Cell[] row =
                [
                    new(record.Description),
                        new(record.Stars),
                        new(record.Title),
                        new(record.Category),
                        new(record.CreatedAt),
                        new(record.UpdatedAt),
                        new(record.Author),
                        new(record.IsActive),
                        new(record.Priority),
                        new(record.Budget),
                        new(record.Identifier.ToString()),
                        new(record.Tags),
                        new(record.Comments),
                        new(record.Rating),
                        new(record.Views)
                ];

                await spreadsheet.AddRowAsync(row, cancellationToken).ConfigureAwait(false);

                totalItems++;
            }

            await spreadsheet.FinishAsync(cancellationToken);

            stream.Position = 0;



            // await s3Client.UploadObjectFromStreamAsync(
            //        bucketName: "s3-bucket-local",
            //        objectKey: fileName,
            //        stream: stream,
            //        additionalProperties: null,
            //        cancellationToken: cancellationToken
            //    );

            await using FileStream uploadStream = new(
          filePath,
          FileMode.Open,
          FileAccess.Read,
          FileShare.Read,
          bufferSize: DefaultBufferSize,
          useAsync: true
     );

            await s3Client.UploadObjectFromStreamAsync(
             bucketName: "s3-bucket-local",
             objectKey: fileName,
             stream: uploadStream,
             additionalProperties: null,
             cancellationToken: cancellationToken
         );
        }
        catch (Exception ex)
        {
            System.Console.WriteLine("Error " + ex.Message);
            throw;
        }
        finally
        {
            DeleteTemporaryFile(filePath);
        }
    }

    private static void DeleteTemporaryFile(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch (Exception ex)
        {
            System.Console.WriteLine("Error on attempt to delete file " + ex.Message);
        }
    }
}
