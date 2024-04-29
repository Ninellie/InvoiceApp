using System.ComponentModel.DataAnnotations;
using InvoiceApp;
using Microsoft.Extensions.Options;
using Notion.Client;
using Page = Notion.Client.Page;

namespace Service;

public class NotionOptions
{
    [Required]
    public string BaseUrl { get; set; }

    [Required]
    public string NotionVersion { get; set; }

    [Required]
    public string AuthToken { get; set; }
}

public class InvoiceDataService
{
    private readonly NotionClient _notionClient;

    #region Constructor
    public InvoiceDataService(IOptions<NotionOptions> options)
    {
        _notionClient = NotionClientFactory.Create(new ClientOptions
        {
            AuthToken = options.Value.AuthToken,
            BaseUrl = options.Value.BaseUrl,
            NotionVersion = options.Value.NotionVersion
        });
    }
    public InvoiceDataService(NotionOptions options)
    {
        _notionClient = NotionClientFactory.Create(new ClientOptions
        {
            AuthToken = options.AuthToken,
            BaseUrl = options.BaseUrl,
            NotionVersion = options.NotionVersion
        });
    }

    #endregion
    public InvoiceData GetInvoice(string pageId)
    {
        // Получение страницы накладной
        var invoicePage = GetInvoicePageAsync(pageId).Result;
        var properties = invoicePage.Properties;
        var config = new InvoiceDataPropertySettings();
        var date = GetPropValue(properties[config.Date]).ToString();
        var invoiceId = GetPropValue(properties[config.Number]).ToString();
        var relations = GetPositionsAsync(pageId, properties[config.Positions].Id).Result as ListPropertyItem;

        // todo Вставить куда-нибудь проверку на разные заказы

        // Получение позиций накладной
        var items = new List<ItemData>();

        foreach (var simplePropertyItem in relations.Results)
        {
            var relationPropertyItem = (RelationPropertyItem)simplePropertyItem;
            var page = GetInvoicePageAsync(relationPropertyItem.Relation.Id).Result;
            var idProperty = GetPropValue(page.Properties[config.Id]).ToString();
            var itemId = int.Parse(idProperty);
            var diameter = 0;

            if (GetPropValue(page.Properties[config.Diameter]) is List<string> diameterRelationProp)
            {
                var diameterPage = GetInvoicePageAsync(diameterRelationProp.First()).Result;
                var diameterProperty = GetPropValue(diameterPage.Properties[config.DiameterTitle]).ToString();
                if (diameterProperty != null)
                {
                    diameter = int.Parse(diameterProperty);
                }
            }

            var lengthPerItem = (double)GetPropValue(page.Properties[config.LengthPerItem]);
            var massPerMeter = (double)GetPropValue(page.Properties[config.MassPerMeter]);
            var amount = Convert.ToInt32((double)GetPropValue(page.Properties[config.Amount]));
            var techObject = (string)GetPropValue(page.Properties[config.Object]);

            var item = new ItemData
            {
                id = itemId,
                diameter = diameter,
                lengthPerItem = lengthPerItem,
                Amount = amount,
                MassPerMeter = massPerMeter,
                TechObject = techObject
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

    private static object GetPropValue(PropertyValue p)
    {
        switch (p)
        {
            case TitlePropertyValue titlePropertyValue:
                return titlePropertyValue.Title.FirstOrDefault().PlainText;
            case RichTextPropertyValue richTextPropertyValue:
                return richTextPropertyValue.RichText.FirstOrDefault()?.PlainText;
            case NumberPropertyValue numberPropertyValue:
                return numberPropertyValue.Number.Value;
            case SelectPropertyValue selectPropertyValue:
                return selectPropertyValue.Select.Name;
            case DatePropertyValue datePropertyValue:
                return datePropertyValue.Date.Start.Value.ToShortDateString();
            case RelationPropertyValue relationPropertyValue:
                var relations = relationPropertyValue.Relation;
                var idList = new List<string>(relations.Count);
                idList.AddRange(relations.Select(objectId => objectId.Id));
                return idList;
            case FormulaPropertyValue formulaPropertyValue:
                return formulaPropertyValue.Formula.String;
            case RollupPropertyValue rollupPropertyValue:
                return rollupPropertyValue.Rollup.Number.Value;
                    default:
                return null;
        }
    }

    private async Task<Page> GetInvoicePageAsync(string pageId)
    {
        var invData = await _notionClient.Pages.RetrieveAsync(pageId);
        return invData;
    }

    private async Task<IPropertyItemObject> GetPositionsAsync(string pageId, string propId)
    {
        var parameters = new RetrievePropertyItemParameters()
        {
            PageId = pageId,
            PageSize = null,
            PropertyId = propId,
            StartCursor = ""
        };

        var relations = await _notionClient.Pages.RetrievePagePropertyItemAsync(parameters);
        return relations;
    }
}