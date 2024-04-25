using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.Options;

namespace Service.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class InvoiceController : ControllerBase
    {
        private readonly ILogger<InvoiceController> _logger;
        private readonly NotionOptions _options;

        public InvoiceController(ILogger<InvoiceController> logger, IOptions<NotionOptions> options)
        {
            _logger = logger;
            _options = options.Value;
        }

        [HttpPost(Name = "CreateInvoice")]
        public string Post()
        {
            var invoiceDataService = new InvoiceDataService(_options);
            //var invoiceItemId = $"2e33e7a86e154fb48376347ec8fb09c4";
            var invoiceItemId = $"c4b77a7093a049e7a15aaca1d7406ac0";
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
                    System.IO.File.Delete(invoiceDocument);
                }
            }

            return invoiceDocument;

            //return "OK";
            //return _options;
        }
    }
}
