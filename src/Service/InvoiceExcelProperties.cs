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
    public CellAddress OrderFirstCell { get; } = new(16, 1);
    public CellAddress MassSumFirstCell { get; } = new(19, 8);
    public CellAddress MassPerDiameterFirstCell { get; } = new(19, 3);

    public int IdColumn { get; } = 1;
    public int DiameterColumn { get; } = 3;
    public int LengthPerItemColumn { get; } = 4;
    public int AmountColumn { get; } = 5;
    public int TotalLengthColumn { get; } = 6;
    public int MassPerMeterColumn { get; } = 7;
    public int TotalMass { get; } = 8;
}