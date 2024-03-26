using System.Net.Http.Headers;
using InvoiceApp;
using Notion.Client;

namespace Service;

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
    public InvoiceData GetInvoice(string pageId)
    {
        // Получение страницу накладной
        var invoicePage = GetInvoicePageAsync(pageId).Result;
        var properties = invoicePage.Properties;
        
        var date = GetPropertyValue(properties[$"Дата"]).ToString();
        var invoiceId = GetPropertyValue(properties[$"Номер"]).ToString();
        var relations = GetPropertyValue(properties[$"Позиции"]) as List<string>;

        // Получение позиций накладной
        var items = new List<ItemData>();

        foreach (var id in relations)
        {
            var page = GetInvoicePageAsync(id).Result;
            var idProperty = GetPropertyValue(page.Properties[$"Id"]).ToString();
            var itemId = int.Parse(idProperty);
            var diameterProperty = GetPropertyValue(page.Properties[$"Ø"]).ToString();
            var diameter = int.Parse(diameterProperty);
            var lengthPerItem = (double)GetPropertyValue(page.Properties[$"м/ед"]);
            var amount = Convert.ToInt32((double)GetPropertyValue(page.Properties[$"Ед"]));
            var massPerMeter = (double)GetPropertyValue(page.Properties[$"Кг/м"]);
            var TechObject = (string)GetPropertyValue(page.Properties[$"TechObject"]);

            var item = new ItemData
            {
                id = itemId,
                diameter = diameter,
                lengthPerItem = lengthPerItem,
                amount = amount,
                massPerMeter = massPerMeter,
                TechObject = TechObject
            };
            items.Add(item);
        }

        var invoiceData = new InvoiceData()
        {
            Date = date,
            Id = invoiceId,
            Orders = new List<OrderData>()
        };
        
        // Разбор позиций по заказам
        foreach (var itemData in items)
        {
            var orderExist = false;
            foreach (var invoiceDataOrder in invoiceData.Orders)
            {
                if (invoiceDataOrder.name != itemData.TechObject) continue;
                invoiceDataOrder.items.Add(itemData);
                orderExist = true;
                break;
            }

            if (orderExist)
            {
                continue;
            }

            var addedOrder = new OrderData
            {
                name = itemData.TechObject,
                items = new List<ItemData> { itemData },
            };

            invoiceData.Orders.Add(addedOrder);
        }

        return invoiceData;
    }

    private static object GetPropertyValue(PropertyValue p)
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
            case NumberPropertyValue numberPropertyValue:
                return numberPropertyValue.Number.Value;
            case TitlePropertyValue titlePropertyValue:
                return titlePropertyValue.Title.FirstOrDefault().PlainText;
            case SelectPropertyValue selectPropertyValue:
                return selectPropertyValue.Select.Name;
            case FormulaPropertyValue formulaPropertyValue:
                return formulaPropertyValue.Formula.String;
            default:
                return null;
        }
    }

    private async Task<Page> GetInvoicePageAsync(string pageId)
    {
        var invData = await _notionClient.Pages.RetrieveAsync(pageId);
        return invData;
    }
}