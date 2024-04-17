using Service;

namespace UnitTests
{
    public class UnitTest1
    {
        [Fact]
        public async Task CreatingInvoiceTest()
        {
            var invoiceItemId = $"2e33e7a86e154fb48376347ec8fb09c4";
            var httpClient = new HttpClient();
            var invoiceDataService = new InvoiceDataService(httpClient);
            var invoiceData = invoiceDataService.GetInvoice(invoiceItemId);
            var documentCreator = new InvoiceExcelCreator();
            string? invoiceDocument = null;
            try
            {
                invoiceDocument = documentCreator.CreateInvoice(invoiceData, "Templates/InvoiceTemplate.xlsx");
            }
            finally
            {
                if (invoiceDocument != null)
                {
                    File.Delete(invoiceDocument);
                }
            }
            //Assert.Equal("", invoiceDocument);
        }
    }
}