using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace InvoiceApp;



public class NotionInvoiceData
{
    [JsonProperty("properties")]
    public NotionInvoiceProperties Properties { get; set; }
}

public class NotionInvoiceProperties
{
    [JsonProperty("Номер")] public string Id { get; set; }
    [JsonProperty("Дата")] public string Date { get; set; }
}


public class InvoiceDataService
{
    private const string Url = $"https://api.notion.com/v1/pages/";
    private readonly HttpClient _httpClient;
    public InvoiceDataService(HttpClient httpClient, string token, string notionVersion)
    {
        _httpClient = httpClient;
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        _httpClient.DefaultRequestHeaders.Add($"Notion-Version", notionVersion);
        //_httpClient.DefaultRequestHeaders.Add($"Notion-Version", $"2022-06-28");
    }

    public InvoiceData GetInvoice(string pageId)
    {
        var invoiceData = GetInvoicePropertiesAsync(pageId).Result;
        var records = new List<ItemData>();
    }

    private async Task<InvoiceData> GetInvoicePropertiesAsync(string pageId)
    {
        using var response = await _httpClient.GetAsync(Url + pageId);
        response.EnsureSuccessStatusCode();
        var content = await response.Content.ReadAsStringAsync();
        var invoice = JsonConvert.DeserializeObject<NotionInvoiceData>(content)
                      ?? throw new InvalidOperationException("Deseralization null");
        return invoice.Properties;
    }

    private async Task<InvoiceData> GetPositionPropertiesAsync(string pageId)
    {
        using var response = await _httpClient.GetAsync(Url + pageId);
        response.EnsureSuccessStatusCode();
        var content = await response.Content.ReadAsStringAsync();
        var invoice = JsonConvert.DeserializeObject<NotionInvoiceData>(content)
                      ?? throw new InvalidOperationException("Deseralization null");
        return invoice.Properties;
    }
}