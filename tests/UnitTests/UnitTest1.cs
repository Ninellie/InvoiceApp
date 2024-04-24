using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Service;

namespace UnitTests
{
    public class UnitTest1
    {
        [Fact]
        public async Task CreatingInvoiceTest()
        {
            var factory = new WebApplicationFactory<Program>();
            var options = factory.Services.GetRequiredService<IOptions<NotionOptions>>();
            var invoiceDataService = new InvoiceDataService(options);
            var invoiceItemId = $"2e33e7a86e154fb48376347ec8fb09c4";
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
        }
    }
}