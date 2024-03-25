namespace InvoiceApp;

public class ItemData
{
    public int id { get; set; } // order in techObject
    public int diameter { get; set; } // millimeters
    public double lengthPerItem { get; set; } // meters
    public int amount { get; set; }
    public double TotalLength => lengthPerItem * amount;
    public double massPerMeter { get; set; } // kg
    public double TotalMass => massPerMeter * lengthPerItem * amount;
    //
    public string TechObject { get; set; }
}