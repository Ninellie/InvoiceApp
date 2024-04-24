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

        [HttpPost(Name = "CreateInvoice")]
        public NotionOptions Post()
        {
            //return "OK";
            return _options;
        }
    }
}
