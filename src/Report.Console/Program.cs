var records = ReportInMemoryRepository.GetRecordsAsync(10_000);

var generator = new ExcelReportGenerator(new AmazonS3Client("test", "test", new AmazonS3Config
{
    ServiceURL = "http://localhost:4566",
    ForcePathStyle = true,
}));

await generator.GenerateAsync(records);
