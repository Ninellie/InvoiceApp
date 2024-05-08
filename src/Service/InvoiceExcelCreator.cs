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

        var props = new InvoiceExcelProperties();

        // Open the copied template workbook. 
        IWorkbook workbook;
        using (FileStream fileStream = new FileStream(templatePath, FileMode.Open, FileAccess.ReadWrite))
        {
            workbook = new XSSFWorkbook(fileStream);
        }

        // Создание нового листа
        var sheet = workbook.CloneSheet(props.TemplateSheet);

        sheet.GetRow(props.NumberCell.Row).GetCell(props.NumberCell.Column).SetCellValue(invoiceData.Id);
        sheet.GetRow(props.DateCell.Row).GetCell(props.DateCell.Column).SetCellValue(invoiceData.Date);
        //sheet.ActiveCell = props.DateCell;

        // Заполнение масс для каждого диаметра
        var massPerDiameters = invoiceData.MassPerDiameters;
        var massSumRow = props.MassSumFirstCell.Row;

        var massAll = massPerDiameters.Values.Sum();

        sheet.GetRow(19).GetCell(8).SetCellValue(massAll);

        foreach (var massPerDiameter in massPerDiameters)
        {
            var rowValue = $"Ø{massPerDiameter.Key} = {(massPerDiameter.Value / 1000).Round(1)} t";
            var massPerDiameterCell = sheet.CreateRow(massSumRow).CreateCell(props.MassPerDiameterFirstCell.Column);
            massPerDiameterCell.SetCellValue(rowValue);
            massSumRow++;
        }

        //creating style
        var itemCellStyle = workbook.CreateCellStyle();
        itemCellStyle.BorderTop = BorderStyle.Thin;
        itemCellStyle.BorderBottom = BorderStyle.Thin;
        itemCellStyle.BorderLeft = BorderStyle.Thin;
        itemCellStyle.BorderRight = BorderStyle.Thin;

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
        for (int i = 0; i < invoiceData.Orders.Count; i++)
        {
            var order = invoiceData.Orders[i];
            var orderNameCell = sheet.CreateRow(orderRowIndex).CreateCell(props.OrderFirstCell.Column);
            orderNameCell.SetCellValue(order.name);
            orderRowIndex++;
            orderNameCell.CellStyle = itemCellStyle;
            orderNameCell.CellStyle.Alignment = HorizontalAlignment.Center;

            foreach (var item in order.items)
            {
                var itemRow = sheet.CreateRow(orderRowIndex);

                itemRow.CreateCell(props.IdColumn).SetCellValue(item.id);
                itemRow.GetCell(props.IdColumn).CellStyle = itemCellStyle;

                itemRow.CreateCell(props.DiameterColumn).SetCellValue(item.diameter);
                itemRow.GetCell(props.DiameterColumn).CellStyle = itemCellStyle;

                itemRow.CreateCell(props.LengthPerItemColumn).SetCellValue(item.lengthPerItem);
                itemRow.GetCell(props.LengthPerItemColumn).CellStyle = itemCellStyle;

                itemRow.CreateCell(props.AmountColumn).SetCellValue(item.Amount);
                itemRow.GetCell(props.AmountColumn).CellStyle = itemCellStyle;

                itemRow.CreateCell(props.TotalLengthColumn).SetCellValue(item.TotalLength);
                itemRow.GetCell(props.TotalLengthColumn).CellStyle = itemCellStyle;

                itemRow.CreateCell(props.MassPerMeterColumn).SetCellValue(item.MassPerMeter);
                itemRow.GetCell(props.MassPerMeterColumn).CellStyle = itemCellStyle;

                itemRow.CreateCell(props.TotalMass).SetCellValue(item.TotalMass);
                itemRow.GetCell(props.TotalMass).CellStyle = itemCellStyle;

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
}