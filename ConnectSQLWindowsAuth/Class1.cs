using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using Newtonsoft.Json.Linq;

namespace ConnectSQLWindowsAuth
{
    public class Class1
    {
        public string connectSQL_GetDataInExcel(string connectionString, string query, string excelFilePath,string configFile)
        {
            SqlConnection cnn;
            SqlDataAdapter da;
            SqlCommand cmd;
            DataSet ds;
            //connetionString = connectionString;//@"Data Source=AAIN1927MWQR\SQLEXPRESS01;Initial Catalog=master;Integrated security=True";
            cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
                Console.WriteLine("Connection Open ! ");
                String queryTrimmed = query.Trim().ToLower();
                System.Data.DataTable dataTable = new System.Data.DataTable();
                if (queryTrimmed.Substring(0, 6) == "select")
                {
                    cmd = new SqlCommand(query, cnn);
                    SqlDataReader dr = cmd.ExecuteReader();

                    string fileName = excelFilePath;//Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "ExcelReport.xlsx";

                  
                    object misValue = System.Reflection.Missing.Value;

                    // Remove the old excel report file
                    try
                    {
                        //FileInfo oldFile = new FileInfo(fileName);
                        //if (oldFile.Exists)
                        //{
                        //    File.SetAttributes(oldFile.FullName, FileAttributes.Normal);
                        //    oldFile.Delete();
                        //}
                    }
                    catch (Exception ex)
                    {
                      
                    }

                    try
                    {
                        //Excel.Application xlsApp;
                        //Excel.Workbook xlsWorkbook;
                        //Excel.Worksheet xlsWorksheet;
                        Excel.Application xlsApp = new Excel.Application();
                        Excel.Workbook xlsWorkbook = xlsApp.Workbooks.Add(misValue);
                        Excel.Worksheet xlsWorksheet = (Excel.Worksheet)xlsWorkbook.Sheets[1];

                        int i = 1;

                        if (dr.HasRows)
                        {
                            for (int j = 0; j < dr.FieldCount; ++j)
                            {
                                xlsWorksheet.Cells[i, j + 1] = dr.GetName(j);
                              //  return dr.GetName(j);
                            }
                            ++i;
                        }

                        while (dr.Read())
                        {
                            for (int j = 1; j <= dr.FieldCount; ++j)
                            {
                                string Value = Convert.ToString(dr.GetValue(j - 1));
                                string config = File.ReadAllText(configFile);
                                string[] configArray = config.Split(',');
                                if (configArray.Contains(Value.Substring(0, 1)))
                                {
                                    Value = "'" + Value;
                                }
                                xlsWorksheet.Cells[i, j] = Value;
                            }
                            ++i;
                        }

                        Excel.Range range = xlsWorksheet.get_Range("A1", "I" + (i + 2).ToString());
                        range.Columns.AutoFit();

                        xlsWorkbook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue,
                            Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges, misValue, misValue, misValue, misValue);
                        xlsWorkbook.Close(true, misValue, misValue);
                        xlsApp.Quit();

                        //ReleaseObject(xlsWorksheet);
                        //ReleaseObject(xlsWorkbook);
                        //ReleaseObject(xlsApp);


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error creating Excel report: " + ex.Message);
                    }


                }
                else
                {
                    cmd = new SqlCommand();
                    cmd.Connection = cnn;
                    cmd.CommandText = query;
                    cmd.ExecuteNonQuery();
                }
                cnn.Close();
                return "success";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Can not open connection ! "+ex.Message);
                cnn.Close();
                return "success";
            }
        }

        static private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        public  string connectSQL_GetDataInCsv(string connectionString, string query, string csvFilePath)
        {
            //INSERT INTO [Table_2] ([name],[place]) VALUES ('sikha','kerala')
            //UPDATE [Table_2] SET [name]='Sikha P' WHERE [place]='kerala'
            //DELETE [Table_2] WHERE [place]='kerala'
            SqlConnection cnn;
            SqlDataAdapter da;
            SqlCommand cmd;
            DataSet ds;
            //connetionString = connectionString;//@"Data Source=AAIN1927MWQR\SQLEXPRESS01;Initial Catalog=master;Integrated security=True";
            cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
                Console.WriteLine("Connection Open ! ");
                String queryTrimmed = query.Trim().ToLower();

                if (queryTrimmed.Substring(0, 6) == "select")
                { 
                    da = new SqlDataAdapter(query, cnn);
                    ds = new DataSet();
                    da.Fill(ds);
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    if (ds.Tables.Count > 0)
                    {
                        dataTable = ds.Tables[0];
                        ToCSV(dataTable, csvFilePath);
                    }
                }
                else
                {
                    cmd = new SqlCommand();
                    cmd.Connection = cnn;
                    cmd.CommandText = query;
                    cmd.ExecuteNonQuery();
                }
                cnn.Close();
                return "Sucesss";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Can not open connection ! ");
                cnn.Close();
                return "Sucesss"+ ex.Message;
            }
        }

        private  void ToCSV(System.Data.DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers    
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }


      
        }
}
