using ConsoleText;
using Dapper;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Utils.Extensions;

namespace ConsoleTest
{
    class Program
    {
        readonly static ConnectionBase cn = new ConnectionBase();

        #region Query
        readonly static string INSERT_TB_DADOS_CLIENTE = @"
                    INSERT INTO TB_DADOS_CLIENTE
                        VALUES(@nr_cliente, @tx_cpf, @nm_cliente, @dt_nasc)";

        readonly static string SELECT_TB_DADOS_CLIENTE = @"
                    SELECT * FROM TB_DADOS_CLIENTE";
        #endregion


        [STAThread]
        static void Main(string[] args)
        {
            int op = 0;

            Console.WriteLine("----Choose----");
            Console.WriteLine("1- To Import--");
            Console.WriteLine("2- To Export--");
            op = Convert.ToInt32(Console.ReadLine());

            int lines = 0;
            switch (op)
            {
                case 1:
                    //GetCSV and insert into DB
                    lines = ImportaCSV();
                    Console.WriteLine("Affected Lines: " + lines);
                    break;
                case 2:
                    lines = ExportCSV();
                    Console.WriteLine("created files: " + lines);
                    break;
                default:
                    break;
            }


            Console.WriteLine("Press any key to end");
            Console.ReadLine();

        }
        private static int ImportaCSV(/*StreamReader csv*/)
        {
            var pathFileName = new OpenFile().GetStreamSelectedPath();
            StreamReader sr = new StreamReader(File.OpenRead(pathFileName));
            //StreamReader sr = new StreamReader(File.OpenRead("Arquivo_Importacao.csv"));            

            char[] delimiter = new char[] { ';' };
            string[] columnheaders = sr.ReadLine().Split(delimiter);

            string currentLine;
            int countLines = 0;
            // currentLine will be null when the StreamReader reaches the end of file
            while ((currentLine = sr.ReadLine()) != null)
            {
                var splitLine = currentLine.Split(';');

                using (IDbConnection dbConnection = cn.ObterConexao())
                {
                    //var insert = "INSERT INTO TB_DADOS_CLIENTE VALUES(@nr_cliente, @tx_cpf, @nm_cliente, @dt_nasc)";

                    var parameters = new DynamicParameters();
                    parameters.Add("@nr_cliente", splitLine[0].Trim());
                    parameters.Add("@tx_cpf", splitLine[1].Trim());
                    parameters.Add("@nm_cliente", splitLine[2].Trim());
                    parameters.Add("@dt_nasc", splitLine[3].Trim());

                    dbConnection.Query(INSERT_TB_DADOS_CLIENTE, parameters);
                    countLines++;
                }

                //SqlCommand command;
                //    using (SqlConnection connection = new SqlConnection(ConnectionString))
                //    {
                //        connection.ConnectionString = ConnectionString;
                //        connection.Open();

                //        var insert = "INSERT INTO TB_DADOS_CLIENTE VALUES(@nr_cliente, @tx_cpf, @nm_cliente, @dt_nasc)";

                //        command = new SqlCommand(insert, connection);
                //        command.Parameters.AddWithValue("@nr_cliente", splitLine[0].Trim());
                //        command.Parameters.AddWithValue("@tx_cpf", splitLine[1].Trim());
                //        command.Parameters.AddWithValue("@nm_cliente", splitLine[2].Trim());
                //        command.Parameters.AddWithValue("@dt_nasc", splitLine[3].Trim());

                //        command.ExecuteNonQuery();
                //        countLines++;

                //        command.Dispose();
                //        connection.Close();
                //    }
            }

            return countLines;
        }
        private static int ExportCSV()
        {
            try
            {
                int countLines = 0;

                IEnumerable<ExpandoObject> strExcel = GetGenericObjExpando(SELECT_TB_DADOS_CLIENTE);

                string path = Directory.GetCurrentDirectory();
                string target = @"c:\temp";

                if (!Directory.Exists(target))
                {
                    Directory.CreateDirectory(target);
                }
                //set CurrentDirectory
                var currentDirectory = Environment.CurrentDirectory = (target);

                string writeFile = currentDirectory + @"\file.xls";

                File.WriteAllText(writeFile, AssemblyTables(strExcel, "Title"));

                countLines++;

                return countLines;

            }
            catch (Exception e)
            {
                throw;
            }
        }

        private static IEnumerable<ExpandoObject> GetGenericObjExpando(string query, DynamicParameters parameters = null)
        {
            IEnumerable<ExpandoObject> results;

            using (IDbConnection dbConnection = cn.ObterConexao())
            {
                results = dbConnection.Query(query, parameters)
                    .Select(x => (ExpandoObject)ExpandoObjectExtension.ToExpandoObject(x));
            }

            return results;
        }

        private static string AssemblyTables(IEnumerable<ExpandoObject> objList, string cabecalho)
        {
            var header = objList.ToList()[0].Select(p => p.Key);
            List<object> lines = new List<object>();

            foreach (var item in objList.ToList())
            {
                lines.Add(item.Select(p => p.Value).ToList());
            }

            var sb = new StringBuilder();
            sb.Append("<table>");

            //title
            var quantidadeColunas = header.Count();
            sb.Append("<tr>");

            sb.Append("<td colspan=\"" + quantidadeColunas + "\" style=\"border: 1px solid #dee2e6; color: #727376; text-align: center; padding: 10px\"><font face=Open Sans size=3 >" + cabecalho + "</font></td>");
            sb.Append("</tr>");


            //body column
            sb.Append("<tr>");
            foreach (var col in header)
            {
                sb.Append("<td style=\"border: 1px solid #dee2e6; color: #727376; padding: 5px\"><font face=Open Sans size=3>" + col + "</ font></td>");
            }
            sb.Append("</tr>");

            int i = 0;
            foreach (var line in lines)
            {
                string styleCell = "";
                if (i % 2 == 0)
                {
                    styleCell = "style=\"color: #2382AF; background-color: #F5F5FA\"";
                }
                else
                {
                    styleCell = "style=\"color: #2382AF\"";
                }
                var lineCols = ((IEnumerable)line).Cast<object>().ToList();

                sb.Append("<tr>");
                foreach (var col in lineCols)
                {
                    sb.Append($"<td {styleCell}><font face=Open Sans size=" + "14px" + ">" + col + "</font></td>");
                }
                sb.Append("</tr>");

                i++;
            }

            sb.Append("</table>");

            return sb.ToString();
        }

        //private void MakeDataTable()
        //{
        //    var pathFileName = new OpenFile().GetStreamSelectedPath();
        //    StreamReader sr = new StreamReader(File.OpenRead(pathFileName));
        //    //StreamReader sr = new StreamReader(File.OpenRead("Arquivo_Importacao.csv"));

        //    DataTable datatable = new DataTable();
        //    char[] delimiter = new char[] { ';' };

        //    string[] columnheaders = sr.ReadLine().Split(delimiter);
        //    foreach (string columnheader in columnheaders)
        //    {
        //        datatable.Columns.Add(columnheader); // I've added the column headers here.
        //    }

        //    while (sr.Peek() > 0)
        //    {
        //        DataRow datarow = datatable.NewRow();
        //        datarow.ItemArray = sr.ReadLine().Split(delimiter);
        //        datatable.Rows.Add(datarow);
        //    }

        //    foreach (DataRow row in datatable.Rows)
        //    {
        //        Console.WriteLine("----Row No: " + datatable.Rows.IndexOf(row) + "----");

        //        foreach (DataColumn column in datatable.Columns)
        //        {
        //            //check what columns you need
        //            if (column.ColumnName == "Id" ||
        //                column.ColumnName == "CPF" ||
        //                column.ColumnName == "Nascimento")
        //            {
        //                Console.Write(column.ColumnName);
        //                Console.Write(" ");
        //                Console.WriteLine(row[column]);
        //            }
        //        }
        //    }
        //    Console.ReadLine();
        //}

    }
}