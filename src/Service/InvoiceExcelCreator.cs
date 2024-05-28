using InvoiceApp;
using MathNet.Numerics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Service;

public class InvoiceExcelCreator
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="invoiceData"></param>
    /// <param name="templatePath"></param>
    /// <returns>Path of the created file</returns>
    /// <exception cref="FileNotFoundException"></exception>
    public string CreateInvoice(InvoiceData invoiceData, string templatePath)
    {
        // TemplateInvoice.xlsx"
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template not exist", templatePath);

        var generatedFilePath = Path.GetTempFileName();

        // Open the copied template workbook. 
        IWorkbook workbook;
        using (FileStream fileStream = new FileStream(templatePath, FileMode.Open, FileAccess.ReadWrite))
        {
            workbook = new XSSFWorkbook(fileStream);
        }

        var props = new InvoiceExcelProperties(workbook);

        // Создание нового листа
        var sheet = workbook.CloneSheet(props.TemplateSheet);

        sheet.GetRow(props.NumberCell.Row).GetCell(props.NumberCell.Column).SetCellValue(invoiceData.Id);
        sheet.GetRow(props.DateCell.Row).GetCell(props.DateCell.Column).SetCellValue(invoiceData.Date);
        //sheet.ActiveCell = props.DateCell;

        // Заполнение масс для каждого диаметра
        var massPerDiameters = invoiceData.MassPerDiameters;
        var massSumRow = props.MassSumFirstCell.Row;

        var massAll = massPerDiameters.Values.Sum();

        sheet.GetRow(props.MassSumFirstCell.Row).GetCell(props.MassSumFirstCell.Column).SetCellValue(massAll);

        foreach (var massPerDiameter in massPerDiameters)
        {
            var rowValue = $"Ø{massPerDiameter.Key} = {(massPerDiameter.Value / 1000).Round(1)} t";
            var massPerDiameterCell = sheet.CreateRow(massSumRow).CreateCell(props.MassPerDiameterFirstCell.Column);
            massPerDiameterCell.SetCellValue(rowValue);
            massSumRow++;
        }

        // inserting new rows
        var orderRowIndex = props.OrderFirstCell.Row;
        var firstRow = orderRowIndex;
        var newRowsNumber = invoiceData.Orders.Count + invoiceData.Orders.Sum(order => order.items.Count);

        if (newRowsNumber != 0)
        {
            var secondRow = orderRowIndex + newRowsNumber;
            sheet.ShiftRows(firstRow, secondRow, newRowsNumber, true, false);
        }

        // filling cells
        foreach (var order in invoiceData.Orders)
        {
            var orderNameCell = sheet.CreateRow(orderRowIndex).CreateCell(props.OrderFirstCell.Column);
            orderNameCell.SetCellValue(order.name);
            orderRowIndex++;
            var cellStyle = props.ItemCellStyle;
            orderNameCell.CellStyle = cellStyle ?? throw new NullReferenceException("Item cell style is null");
            orderNameCell.CellStyle.Alignment = HorizontalAlignment.Center;

            foreach (var item in order.items)
            {
                var itemRow = sheet.CreateRow(orderRowIndex);
                FillItemDataCell(itemRow, props.IdColumn, item.Id, cellStyle);
                FillItemDataCell(itemRow, props.DiameterColumn, item.Diameter, cellStyle);
                FillItemDataCell(itemRow, props.LengthPerItemColumn, item.LengthPerItem, cellStyle);
                FillItemDataCell(itemRow, props.AmountColumn, item.Amount, cellStyle);
                FillItemDataCell(itemRow, props.TotalLengthColumn, item.TotalLength, cellStyle);
                FillItemDataCell(itemRow, props.MassPerMeterColumn, item.MassPerMeter, cellStyle);
                FillItemDataCell(itemRow, props.TotalMass, item.TotalMassRounded, cellStyle);
                orderRowIndex++;
            }
        }

        // Save Excel document
        using (var fileStream = new FileStream(generatedFilePath, FileMode.Create))
        {
            workbook.Write(fileStream);
        }

        //return path of created file
        return generatedFilePath;
    }

    private void FillItemDataCell(IRow row, int column, double value, ICellStyle cellStyle)
    {
        row.CreateCell(column).SetCellValue(value);
        row.GetCell(column).CellStyle = cellStyle;
    }
}