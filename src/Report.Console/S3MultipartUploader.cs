using Amazon.S3.Model;

namespace Report.Console;

public class S3MultipartUploader(IAmazonS3 s3Client)
{
    public async Task UploadMultipartAsync(string bucketName, string key, MemoryStream stream)
    {
        const int partSize = 5 * 1024 * 1024;

        var initiateRequest = new InitiateMultipartUploadRequest
        {
            BucketName = bucketName,
            Key = key,
        };

        var initResponse = await s3Client.InitiateMultipartUploadAsync(initiateRequest);

        var partETags = new List<PartETag>();

        try
        {
            long filePosition = 0;
            int partNumber = 1;
            while (filePosition < stream.Length)
            {
                int bytesToRead = (int)Math.Min(partSize, stream.Length - filePosition);
                byte[] buffer = new byte[bytesToRead];
                stream.Position = filePosition;
                int read = await stream.ReadAsync(buffer, 0, bytesToRead);

                using var uploadStream = new MemoryStream(buffer, 0, read);

                var uploadRequest = new UploadPartRequest
                {
                    BucketName = bucketName,
                    Key = key,
                    UploadId = initResponse.UploadId,
                    PartNumber = partNumber,
                    InputStream = uploadStream,
                    PartSize = read,
                    IsLastPart = (filePosition + read) >= stream.Length,
                };

                var uploadResponse = await s3Client.UploadPartAsync(uploadRequest);
                partETags.Add(new PartETag(partNumber, uploadResponse.ETag));

                filePosition += read;
                partNumber++;
            }

            var completeRequest = new CompleteMultipartUploadRequest
            {
                BucketName = bucketName,
                Key = key,
                UploadId = initResponse.UploadId,
                PartETags = partETags
            };

            await s3Client.CompleteMultipartUploadAsync(completeRequest);
        }
        catch (Exception)
        {
            await s3Client.AbortMultipartUploadAsync(new AbortMultipartUploadRequest
            {
                BucketName = bucketName,
                Key = key,
                UploadId = initResponse.UploadId
            });
            throw;
        }
    }
}