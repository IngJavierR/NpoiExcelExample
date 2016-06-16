using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using NPOI.HSSF.UserModel;

namespace NpoiExample
{
    public class ExcelReader
    {
        public void Read(string pathFile)
        {
            var fs = new FileStream(pathFile, FileMode.Open, FileAccess.Read);
            using (fs)
            {
                int RowNumber = 0;
                var workbook = new HSSFWorkbook(fs);
                var currentWorksheet = workbook.GetSheetAt(0);
                var connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathFile + ";Extended Properties='Excel 8.0;IMEX=1;HDR=YES'";

                string strSQL =
                    "SELECT *" +
                    " FROM [" + currentWorksheet.SheetName + "$] ";

                //Generate OledbConnection
                OleDbConnection conn = new OleDbConnection(connstr);
                conn.Open();

                DataTable dtExcel = new DataTable(); // set up a datatable for store data from excel file.
                OleDbCommand ocmd = new OleDbCommand(strSQL, conn); //insert command which we type at above
                OleDbDataAdapter dta = new OleDbDataAdapter(strSQL, conn);
                dta.Fill(dtExcel);
                if (dtExcel.Rows.Count < 1) //valify that whether the record is more that 1.
                    return;
                OleDbDataReader odr = ocmd.ExecuteReader();

                while (odr.Read()) //Start read data row by row.
                {
                    string ProductName = (dtExcel.Rows[RowNumber]["PRODUCT_NAME"].ToString());
                }
                conn.Close();
            }
        }
    }
}
