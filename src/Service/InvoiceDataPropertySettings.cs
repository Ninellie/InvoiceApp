namespace Service;

public class InvoiceDataPropertySettings
{
    public string Date { get; }
    public string Number { get; }
    public string Positions { get; }
    public string Id { get; }
    public string Diameter { get; }
    public string DiameterTitle { get; }
    public string LengthPerItem { get; }
    public string Amount { get; }
    public string MassPerMeter { get; }
    public string Object { get; }

    public InvoiceDataPropertySettings()
    {
        Date = "Date";
        Number = "Number";
        Positions = "Positions";
        Id = "Id";
        Diameter = "Ø";
        DiameterTitle = "mm";
        LengthPerItem = "м/ед";
        Amount = "Ед";
        MassPerMeter = "kg/m";
        Object = "TechObject";
    }
}