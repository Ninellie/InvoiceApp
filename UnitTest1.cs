using Newtonsoft.Json;

namespace InvoiceApp
{
    public class InvoiceData
    {
        public ParticipantData Customer { get; set; }
        public ParticipantData Vendor { get; set; }

        public string Id { get; set; }
        public string Date { get; set; }

        public List<OrderData> Orders { get; set; }

        public Dictionary<int, double> MassPerDiameters => GetMassPerDiameters();

        // Подсчёт диаметров
        private Dictionary<int, double> GetMassPerDiameters()
        {
            var dictionary = new Dictionary<int, double>();

            foreach (OrderData order in Orders)
            {
                foreach (ItemData item in order.items)
                {
                    if (dictionary.ContainsKey(item.diameter))
                    {
                        dictionary[item.diameter] += item.TotalMass;
                        continue;
                    }
                    dictionary.Add(item.diameter, item.TotalMass);
                }
            }
            return dictionary;
        }
    }

    public class ParticipantData
    {
        public string name;
        public string address;
        public string place;
        public string pib;
        public string pdv;
    }

    public class OrderData // TechObject Data
    {
        public string name;

        public List<ItemData> items;
    }

    public class ItemData
    {
        [JsonProperty("Id"))] public int id { get; set; } // order in techObject
        [JsonProperty("Ø")] public int diameter { get; set; } // millimeters
        [JsonProperty("м/ед")] public double lengthPerItem { get; set; } // meters
        [JsonProperty("Ед")] public int amount { get; set; }
        public double TotalLength => lengthPerItem * amount;
        [JsonProperty("Кг/м")] public double massPerMeter { get; set; } // kg
        public double TotalMass => massPerMeter * lengthPerItem * amount;
        //
        [JsonProperty("TechObject")] public string TechObject { get; set; }
    }

    public class UnitTest1
    {
        [Fact]
        public async Task Test1()
        {
            var invoiceItemId = $"2e33e7a86e154fb48376347ec8fb09c4";
            var httpClient = new HttpClient();

            var invoiceDataService = new InvoiceDataService(httpClient);
            // Заполнение класса накладной позициями из прочитанного csv файла
            //foreach (var itemData in records)
            //{
            //    var orderExist = false;
            //    foreach (var invoiceDataOrder in invoiceData.orders)
            //    {
            //        if (invoiceDataOrder.name != itemData.TechObject) continue;
            //        invoiceDataOrder.items.Add(itemData);
            //        orderExist = true;
            //        break;
            //    }

            //    if (orderExist)
            //    {
            //        continue;
            //    }

            //    var addedOrder = new OrderData
            //    {
            //        name = itemData.TechObject,
            //        items = new List<ItemData> { itemData },
            //    };

            //    invoiceData.orders.Add(addedOrder);
            //}
        }
    }
}