using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using BusinessLayer;

namespace RainfallAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //1. Read excel file and store information in a datatable
                ReadExcelFile readExcelFile = new ReadExcelFile();
                DataTable dt = readExcelFile.ReadExcel(ConfigurationManager.AppSettings["FileName"], false);

                //2. Get Month, Year, Row number of Header and starting column number of Rainfall data
                string monthYear = ConfigurationManager.AppSettings["MonthYear"];
                string monthToDelete = monthYear.Split('-')[1];
                int month = Convert.ToUInt16(monthYear.Split('-')[1]);
                int year = Convert.ToUInt16(monthYear.Split('-')[0]);
                int numberOfDaysDataAvailable = Convert.ToUInt16(ConfigurationManager.AppSettings["DataAvailableDays"]);
                int columnHeaderRowNumber = Convert.ToUInt16(ConfigurationManager.AppSettings["ColumnHeaderRowNumber"]);
                int rainfallDataColumnNumber = Convert.ToUInt16(ConfigurationManager.AppSettings["RainfallDataStartingColumnNumber"]);
                string userId = ConfigurationManager.AppSettings["userId"];

                //3. Create datatable with the excel data
                CreateDataTable createDataTable = new CreateDataTable();
                DataTable finalTable = createDataTable.Create();

                for (int rowIndex = columnHeaderRowNumber; rowIndex < dt.Rows.Count; rowIndex++)
                {
                    for (int columnIndex = rainfallDataColumnNumber - 1; columnIndex < rainfallDataColumnNumber + numberOfDaysDataAvailable - 1; columnIndex++)
                    {
                        DateTime rainfallDate = new DateTime(year, month, Convert.ToUInt16(dt.Rows[columnHeaderRowNumber - 1][columnIndex].ToString()));
                        finalTable.Rows.Add(new Object[] { rainfallDate, DateTime.Now, userId, dt.Rows[rowIndex][rainfallDataColumnNumber - 3], dt.Rows[rowIndex][rainfallDataColumnNumber - 2], (dt.Rows[rowIndex][columnIndex].ToString() == "" ? 0 : dt.Rows[rowIndex][columnIndex]) });
                    }
                }

                //4. Delete old data for the same period and insert new data
                string connection = ConfigurationManager.ConnectionStrings["RainfallConnectionString"].ConnectionString;
                SqlConnection sqlConnection = new SqlConnection(connection);
                SqlBulkCopy sqlBulkObject = new SqlBulkCopy(sqlConnection);
                sqlBulkObject.DestinationTableName = "Actual_Rainfall_Block";
                sqlBulkObject.ColumnMappings.Add("ReportedDate", "Recorded_Date");
                sqlBulkObject.ColumnMappings.Add("CreatedDate", "Created_Date");
                sqlBulkObject.ColumnMappings.Add("CreatedBy", "Created_By");
                sqlBulkObject.ColumnMappings.Add("DistrictName", "District_Name");
                sqlBulkObject.ColumnMappings.Add("BlockName", "Block_Name");
                sqlBulkObject.ColumnMappings.Add("ActualRainfall", "Actual_Rainfall");

                sqlConnection.Open();
                string deleteCommand = $"DELETE FROM Actual_Rainfall_Block WHERE [Recorded_Date] LIKE '{year}-{monthToDelete}%'";
                SqlCommand sqlCommand = new SqlCommand(deleteCommand, sqlConnection);
                sqlCommand.ExecuteNonQuery();
                sqlBulkObject.WriteToServer(finalTable);
                sqlConnection.Close();
                Console.WriteLine("Rainfall data uploaded successfully.");
             }
            catch(Exception ex)
            {
                Console.WriteLine("Error while processing rainfall data: " + ex.Message);
            }
            finally
            {
                Console.ReadLine();
            }
        }
    }
}
