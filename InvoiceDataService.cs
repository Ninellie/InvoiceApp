using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace InvoiceApp;

public class NotionInvoiceData
{
    [JsonProperty("properties")] public NotionInvoiceProperties Properties { get; set; }
}

public class NotionInvoiceProperties
{
    [JsonProperty("Номер")] public NotionRichText Id { get; set; }
    [JsonProperty("Дата")] public NotionDate Date { get; set; }
    [JsonProperty("Позиции")] public NotionRelationList ItemList { get; set; }
}

public class NotionRichText
{
    [JsonProperty("plain_text")] public string Text { get; set; }
}

public class NotionDate
{
    [JsonProperty("start")] public string Start { get; set; }
    [JsonProperty("end")] public string End { get; set; }
}

public class NotionRelationList
{
    [JsonProperty("relation")] public NotionRelation[] NotionRelation { get; set; }
}

public class NotionRelation
{
    [JsonProperty("id")] public string PageId { get; set; }
}

public class NotionItemData
{
    [JsonProperty("properties")]
    public ItemData Properties { get; set; }
}

public class InvoiceDataService
{
    private const string Url = $"https://api.notion.com/v1/pages/";
    private readonly HttpClient _httpClient;

    #region Constructor
    public InvoiceDataService(HttpClient httpClient, string token, string notionVersion)
    {
        _httpClient = httpClient;
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        _httpClient.DefaultRequestHeaders.Add($"Notion-Version", notionVersion);
    }

    public InvoiceDataService(HttpClient httpClient)
        : this
        (
            httpClient, 
            Environment.GetEnvironmentVariable("Notion__Token")
            ?? throw new InvalidOperationException("Notion token is missing"),
            $"2022-06-28"
        )
    {
    }

    #endregion

    public InvoiceData GetInvoice(string pageId)
    {
        var invoiceData = GetInvoicePropertiesAsync(pageId).Result;

        var date = invoiceData.Date.Start;
        var invoiceId = invoiceData.Id.Text;
        var notionItemRelations = invoiceData.ItemList;

        //var records = new List<ItemData>(); todo вытянуть также и позиции из накладной


        var invoice = new InvoiceData()
        {
            Date = date,
            Id = invoiceId,
        };
        return invoice;
    }

    private async Task<NotionInvoiceProperties> GetInvoicePropertiesAsync(string pageId)
    {
        using var response = await _httpClient.GetAsync(Url + pageId);
        response.EnsureSuccessStatusCode();
        var content = await response.Content.ReadAsStringAsync();
        var invoice = JsonConvert.DeserializeObject<NotionInvoiceData>(content)
                      ?? throw new InvalidOperationException("Deseralization null");
        return invoice.Properties;
    }

    //private async Task<InvoiceData> GetPositionPropertiesAsync(string pageId)
    //{
    //    using var response = await _httpClient.GetAsync(Url + pageId);
    //    response.EnsureSuccessStatusCode();
    //    var content = await response.Content.ReadAsStringAsync();
    //    var invoice = JsonConvert.DeserializeObject<NotionInvoiceData>(content)
    //                  ?? throw new InvalidOperationException("Deseralization null");
    //    return invoice.Properties;
    //}
}