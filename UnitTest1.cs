using System.Data.Common;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using CsvHelper;
using System.Globalization;
using MathNet.Numerics;
using NPOI.POIFS.Crypt;

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
            var dictionary = new Dictionary<int, double>();

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

    public class OrderData // TechObject Data
    {
        public string name;

        public List<ItemData> items;
    }

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

    public class RawItemData
    {
        public int Id { get; set; }
        public int Diameter { get; set; }
        public double LengthPerItem { get; set; }
        public int Amount { get; set; }
    }

    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            // Reading csv file
            var records = new List<ItemData>();
            using (var reader = new StreamReader(@"C:\Users\apawl\Documents\InvoiceDocs\InvoiceData.csv"))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    var record = new ItemData
                    {
                         id = csv.GetField<int>("Id"),
                         diameter = csv.GetField<int>("Ø"),
                         lengthPerItem = csv.GetField<double>("м/ед"),
                         amount = csv.GetField<int>("Ед"),
                         massPerMeter = csv.GetField<double>("Кг/м"),
                         TechObject = csv.GetField<string>("TechObject"),
                    };
                    records.Add(record);
                }
            }
            
            // Создание класса данных
            var invoiceData = new InvoiceData
            {
                orders = new List<OrderData>()
            };

            // Заполнение класса накладной позициями из прочитанного csv файла
            foreach (var itemData in records)
            {
                var orderExist = false;
                foreach (var invoiceDataOrder in invoiceData.orders)
                {
                    if (invoiceDataOrder.name != itemData.TechObject) continue;
                    invoiceDataOrder.items.Add(itemData);
                    orderExist = true;
                    break;
                }

                if (orderExist)
                {
                    continue;
                }

                var addedOrder = new OrderData
                {
                    name = itemData.TechObject,
                    items = new List<ItemData> { itemData },
                };

                invoiceData.orders.Add(addedOrder);
            }

            // Открытие существующей рабочей книги
            IWorkbook workbook;
            using (FileStream fileStream = new FileStream(@"C:\Users\apawl\Documents\InvoiceDocs\invoice.xlsx", FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            // Создание нового листа
            var sheet = workbook.CloneSheet(0);

            sheet.GetRow(1).GetCell(2).SetCellValue(invoiceData.vendor.name);
            sheet.GetRow(2).GetCell(2).SetCellValue(invoiceData.vendor.address);
            sheet.GetRow(4).GetCell(2).SetCellValue(invoiceData.vendor.place);
            sheet.GetRow(5).GetCell(2).SetCellValue(invoiceData.vendor.pib);
            //sheet.GetRow(3).GetCell(2).SetCellValue(invoiceData.customer. );

            sheet.GetRow(1).GetCell(7).SetCellValue(invoiceData.customer.name);
            sheet.GetRow(2).GetCell(7).SetCellValue(invoiceData.customer.address);
            sheet.GetRow(3).GetCell(7).SetCellValue(invoiceData.customer.place);
            sheet.GetRow(4).GetCell(7).SetCellValue(invoiceData.customer.pib);
            sheet.GetRow(5).GetCell(7).SetCellValue(invoiceData.customer.pdv);

            var orderRowIndex = 16;

            // Заполнение масс для каждого диаметра
            var massPerDiameters = invoiceData.MassPerDiameters;
            var massSumRow = 19;


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

            //var orderCellStyle = workbook.CreateCellStyle();
            //orderCellStyle.BorderTop = BorderStyle.Thin;
            //orderCellStyle.BorderBottom = BorderStyle.Thin;
            //orderCellStyle.BorderLeft = BorderStyle.Thin;
            //orderCellStyle.BorderRight = BorderStyle.Thin;

            // inserting new rows
            var firstRow = orderRowIndex;
            var newRowsNumber = invoiceData.orders.Count + invoiceData.orders.Sum(order => order.items.Count);
            var secondRow = orderRowIndex + newRowsNumber;
            sheet.ShiftRows(firstRow, secondRow, newRowsNumber, true, false);

            // filling cells
            for (int i = 0; i < invoiceData.orders.Count; i++)
            {
                var order = invoiceData.orders[i];
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
            using (FileStream fileStream = new FileStream(@"C:\Users\apawl\Documents\InvoiceDocs\invoice.xlsx", FileMode.Create))
            {
                workbook.Write(fileStream, false);
            }
        }
    }
}