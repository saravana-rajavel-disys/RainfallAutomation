using System;
using System.Data;

namespace BusinessLayer
{
    public class CreateDataTable
    {
        public DataTable Create()
        {
            DataTable finalTable = new DataTable();
            DataColumn reportedDate = new DataColumn("ReportedDate", typeof(DateTime));
            DataColumn createdDate = new DataColumn("CreatedDate", typeof(DateTime));
            DataColumn createdBy = new DataColumn("CreatedBy", typeof(string));
            DataColumn districtName = new DataColumn("DistrictName", typeof(string));
            DataColumn blockName = new DataColumn("BlockName", typeof(string));
            DataColumn actualRainFall = new DataColumn("ActualRainfall", typeof(decimal));
            finalTable.Columns.Add(reportedDate);
            finalTable.Columns.Add(createdDate);
            finalTable.Columns.Add(createdBy);
            finalTable.Columns.Add(districtName);
            finalTable.Columns.Add(blockName);
            finalTable.Columns.Add(actualRainFall);

            return finalTable;
        }
    }
}
