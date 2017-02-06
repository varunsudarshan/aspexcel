using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebApplication1.Models;
using System.Data.OleDb;
using System.IO;
using System.Data;

namespace WebApplication1.Common
{
    public class ReadExcel
    {

        public List<vorstatus> readfile( string path)
        {
            //string path = @"D:\Employee.xls";
            IList<vorstatus> objExcelCon = ReadExcelfile(path);
            return objExcelCon.ToList();
        }

        public IList<vorstatus> ReadExcelfile(string path)
        {
            IList <vorstatus> objVorInfo = new List<vorstatus>();
            try
            {
                OleDbConnection oledbConn = OpenConnection(path);
                if (oledbConn.State == ConnectionState.Open)
                {
                    objVorInfo = ExtractVorExcel(oledbConn);
                    oledbConn.Close();
                }
            }
            catch (Exception ex)
            {
                // Error
            }
            return objVorInfo;
        }

        private IList<vorstatus> ExtractVorExcel(OleDbConnection oledbConn)
        {
            OleDbCommand cmd = new OleDbCommand(); ;
            OleDbDataAdapter oleda = new OleDbDataAdapter();
            DataSet dsVorInfo = new DataSet();

            cmd.Connection = oledbConn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [Vor$]"; //Excel Sheet Name ( Vor )
            oleda = new OleDbDataAdapter(cmd);
            oleda.Fill(dsVorInfo, "Vor");

            var dsVorInfoList = dsVorInfo.Tables[0].AsEnumerable().Select(s => new vorstatus
            {
                Region = Convert.ToString(s["Region"]!= DBNull.Value? s["Region"] : ""),
                Segment = Convert.ToString(s["Segment"] != DBNull.Value? s["Segment"] : "")
           

            }).ToList();

            return dsVorInfoList;
        }


        private OleDbConnection OpenConnection(string path)
        {
            OleDbConnection  oledbConn = null;
            try
            {
                if (Path.GetExtension(path) == ".xls")
                    oledbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + path +"; Extended Properties= \"Excel 8.0;HDR=Yes;IMEX=2\"");
                else if (Path.GetExtension(path) == ".xlsx")
                    oledbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" +path + "; Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");

                oledbConn.Open();
            }
            catch (Exception ex)
            {
                //Error
            }
            return oledbConn;
        }


    }
}