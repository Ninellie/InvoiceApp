namespace InvoiceApp;

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
                if (dictionary.ContainsKey(item.Diameter))
                {
                    dictionary[item.Diameter] += item.TotalMass;
                    continue;
                }
                dictionary.Add(item.Diameter, item.TotalMass);
            }
        }
        return dictionary;
    }
}