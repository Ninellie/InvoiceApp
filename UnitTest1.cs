using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;

namespace InvoiceApp
{
    public class InvoiceData
    {
        public ParticipantData customer = new()
        {
            name = "customer name",
            address = "customer address",
            place = "customer place",
            pib = "customer pib",
            pdv = "customer pdv"
        };

        public ParticipantData vendor = new()
        {
            name = "vendor name",
            address = "vendor address",
            place = "vendor place",
            pib = "vendor pib",
            pdv = "vendor pdv"
        };

        public string id;
        public string date;
        public List<OrderData> orders;

        public Dictionary<int, double> MassPerDiameters => GetMassPerDiameters();

        // Подсчёт диаметров
        private Dictionary<int, double> GetMassPerDiameters()
        {
            Dictionary<int, double> dictionary = new Dictionary<int, double>();
            foreach (OrderData order in orders)
            {
                foreach (ItemData item in order.items)
                {
                    if (dictionary.ContainsKey(item.diameter))
                    {
                        dictionary[item.diameter] += item.TotalMass;
                        continue;
                    }
                    dictionary.Add(item.diameter, item.TotalMass);
                }
            }

            return dictionary;
        }
    }

    public class ParticipantData
    {
        public string name;
        public string address;
        public string place;
        public string pib;
        public string pdv;
    }

    public class OrderData
    {
        public string name;

        public List<ItemData> items;
    }

    public class ItemData
    {
        public int id;
        public int diameter;
        public double lengthPerItem;
        public int amount;
        public double TotalLength => lengthPerItem * amount;
        public double massPerMeter;
        public double TotalMass => massPerMeter * lengthPerItem * amount;
    }

    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            // Открытие существующей рабочей книги
            IWorkbook workbook;
            using (FileStream fileStream = new FileStream(@"C:\Users\apawl\Documents\InvoiceDocs\invoice.xlsx", FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            // Создание нового листа
            var sheet = workbook.CloneSheet(0);

            var invoiceData = new InvoiceData();

            sheet.GetRow(1).GetCell(2).SetCellValue(invoiceData.vendor.name);
            sheet.GetRow(2).GetCell(2).SetCellValue(invoiceData.vendor.address);
            sheet.GetRow(4).GetCell(2).SetCellValue(invoiceData.vendor.place);
            sheet.GetRow(5).GetCell(2).SetCellValue(invoiceData.vendor.pib);
            //sheet.GetRow(3).GetCell(2).SetCellValue(invoiceData.customer. );

            sheet.GetRow(1).GetCell(6).SetCellValue(invoiceData.customer.name);
            sheet.GetRow(2).GetCell(6).SetCellValue(invoiceData.customer.address);
            sheet.GetRow(3).GetCell(6).SetCellValue(invoiceData.customer.place);
            sheet.GetRow(4).GetCell(6).SetCellValue(invoiceData.customer.pib);
            sheet.GetRow(5).GetCell(6).SetCellValue(invoiceData.customer.pdv);

            var orderData = new OrderData
            {
                name = new string("Very easy order")
            };
            var items = new List<ItemData>();

            var itemsNumber = 5;
            for (int i = 0; i < itemsNumber; i++)
            {
                var nextItem = new ItemData
                {
                    amount = 10,
                    diameter = 10,
                    id = 1,
                    lengthPerItem = 13.4,
                };

                items.Add(nextItem);
            }

            orderData.items = items;
            invoiceData.orders = new List<OrderData> { orderData };

            var orderRowIndex = 16;
            var itemRowIndex = orderRowIndex + 1;


            // Подсчёт масс для каждого диаметра
            var massPerDiameters = invoiceData.MassPerDiameters;
            var keys = massPerDiameters.Keys;
            var massSumRow = 19;
            foreach (var value in keys.Select(key => $"{key} = {massPerDiameters[key]} kg"))
            {
                sheet.GetRow(massSumRow).GetCell(3).SetCellValue(value);
                massSumRow++;
            }

            //creating style
            var cellStyle = workbook.CreateCellStyle();
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;

            // inserting new rows
            var firstRow = orderRowIndex;
            var newRowsNumber = invoiceData.orders.Count + invoiceData.orders.Sum(order => order.items.Count);
            var secondRow = orderRowIndex + newRowsNumber; // Вместо 100 вставить кол-во всех позиций
            sheet.ShiftRows(firstRow, secondRow, newRowsNumber, true, false);




            // filling cells
            for (int i = 0; i < invoiceData.orders.Count; i++)
            {
                var order = invoiceData.orders[i];
                var orderNameCell = sheet.CreateRow(orderRowIndex).CreateCell(1);
                orderNameCell.SetCellValue(order.name);
                orderRowIndex++;
                orderNameCell.CellStyle = cellStyle;
                orderNameCell.CellStyle.Alignment = HorizontalAlignment.Center;

                for (var j = 0; j < order.items.Count; j++)
                {
                    var itemRow = sheet.CreateRow(itemRowIndex);
                    var item = order.items[j];

                    itemRow.CreateCell(1).SetCellValue(item.id);
                    itemRow.GetCell(1).CellStyle = cellStyle;

                    itemRow.CreateCell(2).SetCellValue(item.diameter);
                    itemRow.GetCell(2).CellStyle = cellStyle;

                    itemRow.CreateCell(3).SetCellValue(item.lengthPerItem);
                    itemRow.GetCell(3).CellStyle = cellStyle;

                    itemRow.CreateCell(4).SetCellValue(item.amount);
                    itemRow.GetCell(4).CellStyle = cellStyle;

                    itemRow.CreateCell(5).SetCellValue(item.TotalLength);
                    itemRow.GetCell(5).CellStyle = cellStyle;

                    itemRow.CreateCell(6).SetCellValue(item.massPerMeter);
                    itemRow.GetCell(6).CellStyle = cellStyle;
                    
                    itemRow.CreateCell(7).SetCellValue(item.TotalMass);
                    itemRow.GetCell(7).CellStyle = cellStyle;

                    itemRowIndex++;
                }

                orderRowIndex += order.items.Count + 1;
            }

            // Сохранение документа Excel
            using (FileStream fileStream = new FileStream(@"C:\Users\apawl\Documents\InvoiceDocs\invoice.xlsx", FileMode.Create))
            {
                workbook.Write(fileStream, false);
            }
        }
    }
}