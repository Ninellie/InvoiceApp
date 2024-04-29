using Microsoft.AspNetCore.Mvc;
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

        [HttpGet(Name = "CreateInvoice")]
        public async Task<IActionResult> GetInvoice()
        {
            var invoiceDataService = new InvoiceDataService(_options);
            var invoiceItemId = "c4b77a7093a049e7a15aaca1d7406ac0";
            var invoiceData = invoiceDataService.GetInvoice(invoiceItemId);
            var documentCreator = new InvoiceExcelCreator();
            var invoiceDocumentPath = documentCreator.CreateInvoice(invoiceData, "Templates/InvoiceTemplate.xlsx");
            var fileBytes = await System.IO.File.ReadAllBytesAsync(invoiceDocumentPath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Invoice.xlsx");


            Response.Headers.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            using var fileStream = new FileStream(invoiceDocumentPath, FileMode.Open);
            await fileStream.CopyToAsync(Response.Body);
            //System.IO.ile.Delete(invoiceDocument);
            return Ok();
        }
    }
}
