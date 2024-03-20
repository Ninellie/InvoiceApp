using System.Net.Http.Headers;
using Newtonsoft.Json;
using Notion.Client;

namespace InvoiceApp;

public class NotionInvoiceData
{
    [JsonProperty("properties")] public NotionInvoiceProperties Properties { get; set; }
}

public class NotionInvoiceProperties
{
    [JsonProperty("Номер")] public NotionRichText Id { get; set; }
    [JsonProperty("Дата")] public NotionDateObject DateObject { get; set; }
    [JsonProperty("Позиции")] public NotionRelationList ItemList { get; set; }
}

public class NotionRichText
{
    [JsonProperty("rich_text")] public NotionRichTextProperties[] RichText { get; set; }
}

public class NotionRichTextProperties
{
    [JsonProperty("plain_text")] public string Text { get; set; }
}

public class NotionDateObject
{
    [JsonProperty("date")] public NotionDate Date { get; set; }
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
    private readonly NotionClient _notionClient;

    #region Constructor
    public InvoiceDataService(HttpClient httpClient, string token, string notionVersion)
    {
        _httpClient = httpClient;
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        _httpClient.DefaultRequestHeaders.Add($"Notion-Version", notionVersion);

        _notionClient = NotionClientFactory.Create(new ClientOptions
        {
            AuthToken = token,
            BaseUrl = Url,
            NotionVersion = notionVersion
        });
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
    private object GetPropertyValue(PropertyValue p)
    {
        switch (p)
        {
            case RichTextPropertyValue richTextPropertyValue:
                return richTextPropertyValue.RichText.FirstOrDefault()?.PlainText;
            case DatePropertyValue datePropertyValue:
                return datePropertyValue.Date.Start.Value.ToShortDateString();
            case RelationPropertyValue relationPropertyValue:
                var relations = relationPropertyValue.Relation;
                var idList = new List<string>(relations.Count);
                idList.AddRange(relations.Select(objectId => objectId.Id));
                return idList;
            default:
                return null;
        }
    }

    public InvoiceData GetInvoice(string pageId)
    {
        // Получение страницу накладной
        var invoiceData = GetInvoicePageAsync(pageId).Result;
        var properties = invoiceData.Properties;
        
        var date = GetPropertyValue(properties[$"Дата"]).ToString();
        var invoiceId = GetPropertyValue(properties[$"Номер"]).ToString();
        var relations = GetPropertyValue(properties[$"Позиции"]) as List<string>;

        // Получение позиций накладной
        var items = new List<ItemData>();
        foreach (var id in relations)
        {
            var page = GetInvoicePageAsync(id).Result;



            var item = new ItemData();
        }

        var invoice = new InvoiceData()
        {
            Date = date,
            Id = invoiceId,
        };
        return invoice;
    }

    private async Task<Page> GetInvoicePageAsync(string pageId)
    {
        var invData = await _notionClient.Pages.RetrieveAsync(pageId);
        return invData;
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