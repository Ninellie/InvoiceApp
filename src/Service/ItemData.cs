namespace InvoiceApp;

public class ItemData
{
    public int id { get; set; } // order in techObject
    public int diameter { get; set; } // millimeters
    public double lengthPerItem { get; set; } // meters
    public int Amount { get; set; }
    public double TotalLength => lengthPerItem * Amount;
    public double MassPerMeter { get; set; } // kg
    public double TotalMass => MassPerMeter * lengthPerItem * Amount;
    //
    public string TechObject { get; set; }
}