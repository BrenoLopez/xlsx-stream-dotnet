using Amazon.S3.Model;
using Microsoft.IO;
using SpreadCheetah;

namespace Report.Console;

public class ExcelReportGenerator(IAmazonS3 s3Client)
{
    /// <summary>
    /// Generates xlsx report files
    /// </summary>
    /// <param name="records">Sample records</param>
    /// <returns></returns>
    public async Task GenerateAsync(IAsyncEnumerable<SampleRecord> records)
    {
        var bucketName = "s3-bucket-local";

        var fileName = "report.xlsx";

        var uploadId = await InitializeMultipartUploadAsync(bucketName, fileName);

        try
        {
            var partSize = 5 * 1024 * 1024;

            var partNumber = 1;

            var partETags = new List<PartETag>();

            using var stream = new RecyclableMemoryStreamManager().GetStream();

            using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

            await spreadsheet.StartWorksheetAsync("Sheet 1");

            await foreach (var record in records)
            {
                Cell[] row = [new(record.Description), new(record.Stars)];

                await spreadsheet.AddRowAsync(row);

                if (stream.Length >= partSize)
                {
                    await spreadsheet.FinishAsync();
                    stream.Position = 0;
                    partETags.Add(await UploadPartAsync(s3Client, bucketName, fileName, uploadId, partNumber++, stream));
                    stream.SetLength(0);
                }
            }

            if (stream.Length > 0)
            {
                await spreadsheet.FinishAsync();
                stream.Position = 0;
                partETags.Add(await UploadPartAsync(s3Client, bucketName, fileName, uploadId, partNumber++, stream));
                stream.SetLength(0);
            }

            await CompleteMultipartUploadAsync(bucketName, fileName, uploadId, partETags);

        } catch
        {
            await AbortMultipartUploadAsync(bucketName, fileName, uploadId);
            throw;
        }
    }

    private static async Task<PartETag> UploadPartAsync(
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

    private async Task<string> InitializeMultipartUploadAsync(string bucketName, string fileName)
    {
        var initiateRequest = new InitiateMultipartUploadRequest
        {
            BucketName = bucketName,
            Key = fileName
        };

        var initiateResponse = await s3Client.InitiateMultipartUploadAsync(initiateRequest);

        return initiateResponse.UploadId;
    }

    private async Task CompleteMultipartUploadAsync(string bucketName, string fileName, string uploadId, List<PartETag> parts)
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

    private async Task AbortMultipartUploadAsync(string bucketName, string fileName, string uploadId)
    {
        var abortRequest = new AbortMultipartUploadRequest
        {
            BucketName = bucketName,
            Key = fileName,
            UploadId = uploadId
        };

        await s3Client.AbortMultipartUploadAsync(abortRequest);
    }
}
