using MathNet.Numerics;

namespace InvoiceApp;

public class ItemData
{
    public int Id { get; set; } // order in techObject
    public int Diameter { get; set; } // millimeters
    public double LengthPerItem { get; set; } // meters
    public int Amount { get; set; }
    public double TotalLength => LengthPerItem * Amount;
    public double MassPerMeter
    {
        get => _massPerMeter.Round(1);
        set => _massPerMeter = value;
    } // kg
    public double TotalMassRounded => (_massPerMeter * LengthPerItem * Amount).Round(1);
    public double TotalMass => _massPerMeter * LengthPerItem * Amount;
    public string TechObject { get; set; }

    private double _massPerMeter;
}