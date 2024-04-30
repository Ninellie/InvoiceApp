using NPOI.SS.Util;

namespace Service;

/// <summary>
/// Rules defining how to fill out the excel table
/// </summary>
public class InvoiceExcelProperties
{
    public int TemplateSheet { get; } = 0;
    public CellAddress NumberCell { get; } = new(9, 4);
    public CellAddress DateCell { get; } = new(11, 4);
    public CellAddress OrderFirstCell { get; } = new(16, 4);
    public CellAddress MassAllFirstCell { get; } = new(19, 8);
    public CellAddress MassPerDiameterFirstCell { get; } = new(19, 3);

    public int IdColumn = 1;
    public int DiameterColumn = 3;
    public int LengthPerItemColumn = 4;
    public int AmountColumn = 5;
    public int TotalLengthColumn = 6;
    public int MassPerMeterColumn = 7;
    public int TotalMass = 8;

}