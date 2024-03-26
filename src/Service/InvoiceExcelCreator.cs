using InvoiceApp;
using MathNet.Numerics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Service;

public class InvoiceExcelCreator
{
    private string _defaultDocumentPath;

    #region Constructor

    public InvoiceExcelCreator(string defaultDocumentPath)
    {
        _defaultDocumentPath = defaultDocumentPath;
    }

    public InvoiceExcelCreator() : this(@"C:\Users\apawl\Documents\InvoiceDocs\invoice.xlsx")
    {}

    #endregion

    public void CreateInvoice(InvoiceData invoiceData)
    {
        // Открытие существующей рабочей книги
        IWorkbook workbook;
        using (FileStream fileStream = new FileStream(_defaultDocumentPath, FileMode.Open, FileAccess.ReadWrite))
        {
            workbook = new XSSFWorkbook(fileStream);
        }

        // Создание нового листа
        var sheet = workbook.CloneSheet(0);


        var invoiceNumberRowIndex = 9;
        var invoiceNumberCellIndex = 4;
        sheet.GetRow(invoiceNumberRowIndex).
            GetCell(invoiceNumberCellIndex).
            SetCellValue(invoiceData.Id);

        var invoiceDateRowIndex = 11;
        var invoiceDateCellIndex = 4;
        sheet.GetRow(11).
            GetCell(invoiceNumberCellIndex).
            SetCellValue(invoiceData.Date);

        var orderRowIndex = 16; // todo удалить магические числа

        // Заполнение масс для каждого диаметра
        var massPerDiameters = invoiceData.MassPerDiameters;
        var massSumRow = 19; // todo удалить магические числа

        var massAll = massPerDiameters.Values.Sum();

        sheet.GetRow(19).GetCell(8).SetCellValue(massAll);

        foreach (var massPerDiameter in massPerDiameters)
        {
            var rowValue = $"Ø{massPerDiameter.Key} = {massPerDiameter.Value.Round(1)} kg";
            sheet.GetRow(massSumRow).GetCell(3).SetCellValue(rowValue);
            massSumRow++;
        }

        //creating style
        var itemCellStyle = workbook.CreateCellStyle();
        itemCellStyle.BorderTop = BorderStyle.Thin;
        itemCellStyle.BorderBottom = BorderStyle.Thin;
        itemCellStyle.BorderLeft = BorderStyle.Thin;
        itemCellStyle.BorderRight = BorderStyle.Thin;

        // inserting new rows
        var firstRow = orderRowIndex;
        var newRowsNumber = invoiceData.Orders.Count + invoiceData.Orders.Sum(order => order.items.Count);
        var secondRow = orderRowIndex + newRowsNumber;
        sheet.ShiftRows(firstRow, secondRow, newRowsNumber, true, false);

        // filling cells
        for (int i = 0; i < invoiceData.Orders.Count; i++)
        {
            var order = invoiceData.Orders[i];
            var orderNameCell = sheet.CreateRow(orderRowIndex).CreateCell(1);
            orderNameCell.SetCellValue(order.name);
            orderRowIndex++;
            orderNameCell.CellStyle = itemCellStyle;
            orderNameCell.CellStyle.Alignment = HorizontalAlignment.Center;

            foreach (var item in order.items)
            {
                var itemRow = sheet.CreateRow(orderRowIndex);

                itemRow.CreateCell(1).SetCellValue(item.id);
                itemRow.GetCell(1).CellStyle = itemCellStyle;

                itemRow.CreateCell(3).SetCellValue(item.diameter);
                itemRow.GetCell(3).CellStyle = itemCellStyle;

                itemRow.CreateCell(4).SetCellValue(item.lengthPerItem);
                itemRow.GetCell(4).CellStyle = itemCellStyle;

                itemRow.CreateCell(5).SetCellValue(item.amount);
                itemRow.GetCell(5).CellStyle = itemCellStyle;

                itemRow.CreateCell(6).SetCellValue(item.TotalLength);
                itemRow.GetCell(6).CellStyle = itemCellStyle;

                itemRow.CreateCell(7).SetCellValue(item.massPerMeter);
                itemRow.GetCell(7).CellStyle = itemCellStyle;

                itemRow.CreateCell(8).SetCellValue(item.TotalMass);
                itemRow.GetCell(8).CellStyle = itemCellStyle;

                orderRowIndex++;
            }
        }

        // Сохранение документа Excel
        using (var fileStream = new FileStream(@"C:\Users\apawl\Documents\InvoiceDocs\invoice.xlsx", FileMode.Create))
        {
            workbook.Write(fileStream);
        }

        //return path of created file
    }
}