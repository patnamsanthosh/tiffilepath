using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

namespace TifFilePath
{
    class Program
    {
        static void Main(string[] args)
        {
            //string path = @"\\w1p01-dfsstr\records2\{last3digitsofreqid}\{requestId}-n.tif";
            string path = @"C:\Santhosh\tiffiles\"; //@"\\w1p01-dfsstr\records2\"; // Add directory name here
            string con = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\santh\Downloads\ImagePath.xlsx; Extended Properties = 'Excel 12.0 Xml;HDR=YES;'";
            DataTable dt = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(con))
            {

                OleDbCommand command = new OleDbCommand("select * from [Sheet2$]", connection);
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "select * from [Sheet2$]";
                    comm.Connection = connection;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);

                    }

                }
            }

            DataTable workTable = new DataTable("TIFfiles");

            workTable.Columns.Add("RequestID", typeof(String));
            workTable.Columns.Add("FilePath", typeof(String));

            foreach (DataRow item in dt.AsEnumerable())
            {
                string reqID = item["RequestId"].ToString();
                string last3Digits = reqID.Substring(reqID.Length - 4, reqID.Length - 1);
                if (Directory.Exists(path + "\\" + last3Digits))
                {
                    string[] files = Directory.GetFiles(path + "\\" + last3Digits);
                    foreach (string filename in files)
                    {
                        if (filename.Contains(reqID))
                        {
                            DataRow newCustomersRow = workTable.NewRow();

                            newCustomersRow["RequestID"] = item["RequestID"];
                            newCustomersRow["FilePath"] = filename;
                            workTable.Rows.Add(newCustomersRow);
                        }
                    }

                }
            }

            if (workTable.Rows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                //adding header
                sb.Append("RequestID,FilePath");
                sb.AppendLine();
                foreach (DataRow dr in workTable.Rows)
                {
                    foreach (DataColumn dc in workTable.Columns)
                        sb.Append(FormatCSV(dr[dc.ColumnName].ToString()) + ",");
                    sb.Remove(sb.Length - 1, 1);
                    sb.AppendLine();
                }
                File.WriteAllText("C:\\Santhosh\\testresults.csv", sb.ToString());
            }
        }

        public static string FormatCSV(string input)
        {
            try
            {
                if (input == null)
                    return string.Empty;

                bool containsQuote = false;
                bool containsComma = false;
                int len = input.Length;
                for (int i = 0; i < len && (containsComma == false || containsQuote == false); i++)
                {
                    char ch = input[i];
                    if (ch == '"')
                        containsQuote = true;
                    else if (ch == ',')
                        containsComma = true;
                }

                if (containsQuote && containsComma)
                    input = input.Replace("\"", "\"\"");

                if (containsComma)
                    return "\"" + input + "\"";
                else
                    return input;
            }
            catch
            {
                throw;
            }
        }
    }
}
