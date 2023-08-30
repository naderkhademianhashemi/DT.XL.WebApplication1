using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ClosedXML;

namespace DT.XL.WebApplication1.page_XL_Xport
{
    public partial class frmXLsx : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        public void DownloadFile(string path)
        {
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("content-disposition", "attachment;filename=" + Path.GetFileName(path));
            Response.WriteFile(path);
            Response.End();
        }

        public static DataTable CreateDT()
        {
            DataTable table = new DataTable("Answers");

            DataColumn column;
            DataRow row;


            // Create new DataColumn, set DataType, ColumnName and add to DataTable.
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "آی دی";
            table.Columns.Add(column);


            // Create second column.
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "آیتم";

            table.Columns.Add(column);

            // Create new DataRow objects and add to DataTable.
            for (int i = 0; i < 10; i++)
            {
                row = table.NewRow();
                row["آی دی"] = i;
                row["آیتم"] = "آیتم " + i.ToString();
                table.Rows.Add(row);
            }
            return table;
        }

        protected void btnXL_Click(object sender, EventArgs e)
        {
            DataTable dt = CreateDT();
            string path = Server.MapPath("~/Excel/");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            using (var wb = new ClosedXML.Excel.XLWorkbook())
            {
                var sheet = wb.Worksheets.Add(dt, "Answers");
                sheet.Style.Font.FontName = "Ayandeh";
                wb.SaveAs(path + "Answers.xlsx");
            }

            DownloadFile(path + "Answers.xlsx");
        }
    }
}