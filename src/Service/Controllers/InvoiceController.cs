using Microsoft.AspNetCore.Mvc;

namespace Service.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class InvoiceController : ControllerBase
    {
        private readonly ILogger<InvoiceController> _logger;

        public InvoiceController(ILogger<InvoiceController> logger)
        {
            _logger = logger;
        }

        [HttpPost(Name = "CreateInvoice")]
        public string Post()
        {
            return "OK";
        }
    }
}
