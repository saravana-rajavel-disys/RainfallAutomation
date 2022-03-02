using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Linq;
using DataTable = System.Data.DataTable;

namespace BusinessLayer
{
    public class ReadExcelFile
    {
        public DataTable ReadExcel(string fileName, bool hasHeader)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(fileName))
                {
                    package.Load(stream);
                }
                var worksheet = package.Workbook.Worksheets.First();

                DataTable table = new DataTable();
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    table.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }

                var startRow = hasHeader ? 2 : 1;
                for (int rowNumber = startRow; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var wsRow = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    DataRow row = table.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return table;
            }
        }
    }
}
