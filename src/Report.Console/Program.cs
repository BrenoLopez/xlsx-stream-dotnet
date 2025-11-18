using Amazon.S3.Model;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Microsoft.IO;

class Program
{
    private static readonly RecyclableMemoryStreamManager recyclableMemoryStreamManager = new();

    static async Task Main(string[] args)
    {
        var s3Client = new AmazonS3Client("test", "test", new AmazonS3Config
        {
            ServiceURL = "http://localhost:4566",
            ForcePathStyle = true,
        });

        var data = ReportInMemoryRepository.GetRecordsAsync(100);

        await GenerateWithOpenXml(s3Client, data);
    }

    public static async Task GenerateWithOpenXml(AmazonS3Client s3Client, IAsyncEnumerable<SampleRecord> data)
    {
        const int partSize = 5 * 1024 * 1024;
        const string bucketName = "s3-bucket-local";
        const string fileName = "teste.xlsx";
        var partETags = new List<PartETag>();
        int partNumber = 1;

        var uploadId = await InitializeMultipartUploadAsync(s3Client, bucketName, fileName);

        try
        {
            using var initialStream = recyclableMemoryStreamManager.GetStream();
            using (var spreadsheet = SpreadsheetDocument.Create(initialStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Data"
                });
            }

            uint rowIndex = 1;
            using var currentPartStream = recyclableMemoryStreamManager.GetStream();

            initialStream.Position = 0;
            await initialStream.CopyToAsync(currentPartStream);
            initialStream.SetLength(0);

            using (var spreadsheet = SpreadsheetDocument.Open(currentPartStream, true))
            {
                var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                await foreach (var item in data)
                {
                    var row = new Row { RowIndex = rowIndex++ };
                    row.Append(
                        new Cell { CellValue = new CellValue(item.Description) },
                        new Cell { CellValue = new CellValue(item.Stars.ToString()) }
                    );
                    sheetData.Append(row);

                    if (currentPartStream.Length >= partSize)
                    {
                        spreadsheet.WorkbookPart.Workbook.Save();
                        partETags.Add(await UploadPartAsync(
                            s3Client, bucketName, fileName, uploadId, partNumber++, currentPartStream));

                        currentPartStream.SetLength(0);
                        initialStream.Position = 0;
                        await initialStream.CopyToAsync(currentPartStream);
                    }
                }

                if (currentPartStream.Length > 0)
                {
                    spreadsheet.WorkbookPart.Workbook.Save();
                    partETags.Add(await UploadPartAsync(
                        s3Client, bucketName, fileName, uploadId, partNumber++, currentPartStream));
                }
            }

            await CompleteMultipartUploadAsync(s3Client, bucketName, fileName, uploadId, partETags);
        }
        catch
        {
            await AbortMultipartUploadAsync(s3Client, bucketName, fileName, uploadId);
            throw;
        }
    }

    private static async Task<string> InitializeMultipartUploadAsync(AmazonS3Client s3Client, string bucketName, string fileName)
    {
        var initiateRequest = new InitiateMultipartUploadRequest
        {
            BucketName = bucketName,
            Key = fileName
        };

        var initiateResponse = await s3Client.InitiateMultipartUploadAsync(initiateRequest);

        return initiateResponse.UploadId;
    }

    private static async Task CompleteMultipartUploadAsync(AmazonS3Client s3Client, string bucketName, string fileName, string uploadId, List<PartETag> parts)
    {

        var completeRequest = new CompleteMultipartUploadRequest
        {
            BucketName = bucketName,
            Key = fileName,
            UploadId = uploadId,
            PartETags = parts
        };
        await s3Client.CompleteMultipartUploadAsync(completeRequest);
    }

    private static async Task AbortMultipartUploadAsync(AmazonS3Client s3Client, string bucketName, string fileName, string uploadId)
    {
        var abortRequest = new AbortMultipartUploadRequest
        {
            BucketName = bucketName,
            Key = fileName,
            UploadId = uploadId
        };
        await s3Client.AbortMultipartUploadAsync(abortRequest);
    }

    public static async Task<PartETag> UploadPartAsync(
    IAmazonS3 s3Client,
    string bucketName,
    string objectKey,
    string uploadId,
    int partNumber,
    Stream stream)
    {
        var uploadRequest = new UploadPartRequest
        {
            BucketName = bucketName,
            Key = objectKey,
            UploadId = uploadId,
            PartNumber = partNumber,
            PartSize = stream.Length,
            InputStream = stream
        };

        var response = await s3Client.UploadPartAsync(uploadRequest);

        return new PartETag
        {
            PartNumber = partNumber,
            ETag = response.ETag
        };
    }
}
