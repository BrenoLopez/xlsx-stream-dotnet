var uploader = new S3MultipartUploader(new AmazonS3Client("test", "test", new AmazonS3Config
{
    ServiceURL = "http://localhost:4566",
    ForcePathStyle = true,
}));

var records = ReportInMemoryRepository.GetRecordsAsync(10_000);

await using var stream = await ExcelReportGenerator.GenerateAsync(records);

await uploader.UploadMultipartAsync("s3-bucket-local", "report.xlsx", stream);
