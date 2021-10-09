using System;
using System.Data;
using System.Data.OleDb;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Wpf_Excel_DataGrid
{
    public class ExcelDados
    {
        public DataView DadosExcel
        {
            get
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook;
                Excel.Worksheet worksheet;
                Excel.Range range;
                //le a planilha da pasta bin/debug
                //workbook = excelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Macoratti.xlsx");
                workbook = excelApp.Workbooks.Open("C:\\Users\\sampa\\source\\repos\\DADOS1.xlsx");
                worksheet = (Excel.Worksheet)workbook.Sheets["Vendas"];

                int column = 0;
                int row = 0;

                range = worksheet.UsedRange;

                DataTable dt = new DataTable();
                
                dt.Columns.Add("Codigo");
                dt.Columns.Add("Nome");
                dt.Columns.Add("Mes");
                dt.Columns.Add("Valor");
                dt.Columns.Add("Valor2");


                for (row = 2; row <= range.Rows.Count; row++)
                {
                    DataRow dr = dt.NewRow();
                    for (column = 1; column <= range.Columns.Count; column++)
                    {
                        dr[column - 1] = (range.Cells[row, column] as Excel.Range).Value2 != null ? (range.Cells[row, column] as Excel.Range).Value2.ToString() : "";
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                workbook.Close(true, Missing.Value, Missing.Value);
                excelApp.Quit();
                return dt.DefaultView;
            }
        }

        public DataTable GetDataTableExcel(string datasource)
        {

            OleDbConnection theConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;data source=" + datasource + ";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1;\"");
            theConnection.Open();
            OleDbDataAdapter theDataAdapter = new OleDbDataAdapter("SELECT * FROM [Vendas$]", theConnection);
            DataTable dt = new DataTable();
            theDataAdapter.Fill(dt);
            return dt;
        }
    }
}
