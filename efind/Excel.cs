using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Data;
namespace efind
{
    public class Excel
    {
        public static List<template> list = new List<template>();
        private static Workbook xlWorkBook = null;
        private static Application xlApp = null;
        private static Worksheet xlWorkSheet = null;
        private static int lastRow = 0;
        private object misValue = System.Reflection.Missing.Value;
        private string connectionString = null;
        private string sql = null;
        private string data = null;
        private int i = 0;
        private  int j = 0; 
        
        //public void Create_logic_for_thread(int count, int no_of_thread)
        //{
            
        //}

        public Excel()
        {
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }
        public void Create_Excel()
        {
            string connetionString = null;
            SqlConnection connection;
            //SqlCommand command;
            //string SQL = "select * from obj as a left join (select a.obj_id , a.parent_id , obj_name , obj_create_time, obj_modify_time , obj_owner ,temp_id , doc_ext, doc_type, doc_version , doc_extpath from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D'" ;
            //SQL = SQL + ") as b on  a.parent_id = b.temp_id where a.parent_id in (select temp_id from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D')";

            string SQL = "select a.obj_id as index_parent_id , a.obj_name as index_name , b.* from obj as a left join (select a.obj_id as docID, a.parent_id , obj_name , obj_create_time, obj_modify_time , obj_owner ,temp_id , doc_ext, doc_type, doc_version , doc_extpath from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D') as b on  a.parent_id = b.temp_id where a.parent_id in (select temp_id from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D') order by docID";
            string password = "x-admin";
            connetionString = "Data Source=RDTEST1\\RECTIDOC;Initial Catalog=RECTIDOC;User ID=sa;Password=" + password;

            connection = new SqlConnection(connetionString);

            SqlDataAdapter dscmd = new SqlDataAdapter(SQL, connection);
            DataSet ds = new DataSet();
            dscmd.Fill(ds);
            string CurrentDocID = null;
            AddHeaderToExcel();
            int count = 0;
            int ExcelCount = 2;
            //index_parent_id [0]
            //obj_name [1]
            //docID [2]
            //parentid[3]
            //obj_name [4]
            //obj_create_time [5]
            //obj_modify_time [6]
            //obj_owner [7]
            //temp_id [8]
            //doc_ext [9]
            //doc_type [10]
            //doc_version [11]
            //doc_extpath [12]


            //for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
            for (i = 0; i <= 5000; i++)
            {

                if (i > ds.Tables[0].Rows.Count)
                    break;

                template temp = new template();
                int currentRow = i;
                if (CurrentDocID != ds.Tables[0].Rows[i].ItemArray[2].ToString())
                {
                    CurrentDocID = ds.Tables[0].Rows[i].ItemArray[2].ToString();
                    count = 0;
                    temp.index_card = "";
                    temp.index1 = ds.Tables[0].Rows[i].ItemArray[1].ToString();
                    temp.Original_Document_Name = ds.Tables[0].Rows[i].ItemArray[4].ToString();
                    temp.Document_Type = ds.Tables[0].Rows[i].ItemArray[10].ToString();
                    temp.version = Int32.Parse(ds.Tables[0].Rows[i].ItemArray[11].ToString());
                    temp.obj_id = Int32.Parse(ds.Tables[0].Rows[i].ItemArray[2].ToString());
                    temp.paraent_id = Int32.Parse(ds.Tables[0].Rows[i].ItemArray[3].ToString());
                    temp.owner = ds.Tables[0].Rows[i].ItemArray[7].ToString();
                    temp.last_modify_date = ds.Tables[0].Rows[i].ItemArray[6].ToString();
                    temp.create_date = ds.Tables[0].Rows[i].ItemArray[5].ToString();
                    temp.dms_path = "";
                    temp.physical_file_path = "";
                }

                for (int p = 1; p < 19; p++)
                {
                    if (CurrentDocID == ds.Tables[0].Rows[currentRow + p].ItemArray[2].ToString())
                    {
                        count++;
                        switch (count)
                        {
                            case 1:
                                temp.index2 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 2:
                                temp.index3 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 3:
                                temp.index4 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 4:
                                temp.index5 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 5:
                                temp.index6 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 6:
                                temp.index7 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 7:
                                temp.index8 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 8:
                                temp.index9 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 9:
                                temp.index10 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 10:
                                temp.index11 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 11:
                                temp.index12 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 12:
                                temp.index13 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 13:
                                temp.index14 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 14:
                                temp.index15 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 15:
                                temp.index16 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 16:
                                temp.index17 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 17:
                                temp.index18 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 18:
                                temp.index19 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                            case 19:
                                temp.index20 = ds.Tables[0].Rows[currentRow + p].ItemArray[1].ToString();
                                break;
                        }
                        i++;
                    }
                    else
                    {
                        p = 20;
                    }
                }

                WriteToExcel(ExcelCount++, temp);
            }


            xlWorkBook.Save();
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            //string SQL3 = "drop table #temp1 ";

            //using (var cmd = _SQLConnection.conn.CreateCommand())
            //{
            //    _SQLConnection.conn.Open();
            //    cmd.CommandText = SQL3;
            //    var result = cmd.ExecuteNonQuery();
            //}


        }


        //public string CheckHasParentFolder(DataSet ds)
        //{

        //    //for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
        //    //{
        //    //    for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
        //}
        public void AddHeaderToExcel()
        {
            try
            {
                xlWorkSheet.Cells[1, 1] ="index_card";
                xlWorkSheet.Cells[1, 2] ="index1";
                xlWorkSheet.Cells[1, 3] ="index2";
                xlWorkSheet.Cells[1, 4] ="index3";
                xlWorkSheet.Cells[1, 5] ="index4";
                xlWorkSheet.Cells[1, 6] ="index5";
                xlWorkSheet.Cells[1, 7] ="index6";
                xlWorkSheet.Cells[1, 8] ="index7";
                xlWorkSheet.Cells[1, 9] ="index8";
                xlWorkSheet.Cells[1, 10] ="index9";
                xlWorkSheet.Cells[1, 11] ="index10";
                xlWorkSheet.Cells[1, 12] ="index11";
                xlWorkSheet.Cells[1, 13] ="index12";
                xlWorkSheet.Cells[1, 14] ="index13";
                xlWorkSheet.Cells[1, 15] ="index14";
                xlWorkSheet.Cells[1, 16] ="index15";
                xlWorkSheet.Cells[1, 17] ="index16";
                xlWorkSheet.Cells[1, 18] ="index17";
                xlWorkSheet.Cells[1, 19] ="index18";
                xlWorkSheet.Cells[1, 20] ="index19";
                xlWorkSheet.Cells[1, 21] ="index20";
                xlWorkSheet.Cells[1, 22] ="Original_Document_Name";
                xlWorkSheet.Cells[1, 23] ="Document_Type";
                xlWorkSheet.Cells[1, 24] ="version";
                xlWorkSheet.Cells[1, 25] ="obj_id";
                xlWorkSheet.Cells[1, 26] ="paraent_id";
                xlWorkSheet.Cells[1, 27] ="owner";
                xlWorkSheet.Cells[1, 28] ="last_modify_date";
                xlWorkSheet.Cells[1, 29] ="create_date";
                xlWorkSheet.Cells[1, 30] ="dms_path";
                xlWorkSheet.Cells[1, 31] ="physical_file_path";
            }
                catch(Exception ex)
            {
                throw ex;
            }
        }

        public void WriteToExcel(int row,  template temp)
        {
            try
            {
                xlWorkSheet.Cells[row, 1] = temp.index_card;
                xlWorkSheet.Cells[row, 2] = temp.index1;
                xlWorkSheet.Cells[row, 3] = temp.index2;
                xlWorkSheet.Cells[row, 4] = temp.index3;
                xlWorkSheet.Cells[row, 5] = temp.index4;
                xlWorkSheet.Cells[row, 6] = temp.index5;
                xlWorkSheet.Cells[row, 7] = temp.index6;
                xlWorkSheet.Cells[row, 8] = temp.index7;
                xlWorkSheet.Cells[row, 9] = temp.index8;
                xlWorkSheet.Cells[row, 10] = temp.index9;
                xlWorkSheet.Cells[row, 11] = temp.index10;
                xlWorkSheet.Cells[row, 12] = temp.index11;
                xlWorkSheet.Cells[row, 13] = temp.index12;
                xlWorkSheet.Cells[row, 14] = temp.index13;
                xlWorkSheet.Cells[row, 15] = temp.index14;
                xlWorkSheet.Cells[row, 16] = temp.index15;
                xlWorkSheet.Cells[row, 17] = temp.index16;
                xlWorkSheet.Cells[row, 18] = temp.index17;
                xlWorkSheet.Cells[row, 19] = temp.index18;
                xlWorkSheet.Cells[row, 20] = temp.index19;
                xlWorkSheet.Cells[row, 21] = temp.index20;
                xlWorkSheet.Cells[row, 22] = temp.Original_Document_Name;
                xlWorkSheet.Cells[row, 23] = temp.Document_Type;
                xlWorkSheet.Cells[row, 24] = temp.version;
                xlWorkSheet.Cells[row, 25] = temp.obj_id;
                xlWorkSheet.Cells[row, 26] = temp.paraent_id;
                xlWorkSheet.Cells[row, 27] = temp.owner;
                xlWorkSheet.Cells[row, 28] = temp.last_modify_date;
                xlWorkSheet.Cells[row, 29] = temp.create_date;
                xlWorkSheet.Cells[row, 30] = temp.dms_path;
                xlWorkSheet.Cells[row, 31] = temp.physical_file_path;
                //xlWorkBook.Save();
            }
            catch (Exception ex)
            {
                throw ex; 
            }

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }

        public int GetNumberOfRecord(string table_name)
        {
            try
            {
                sql = string.Format("select * from {0}", table_name);
                using (SqlDataAdapter dscmd = new SqlDataAdapter(sql, _SQLConnection.conn))
                {
                    DataSet ds = new DataSet();
                    dscmd.Fill(ds);

                    return ds.Tables[0].Rows.Count;
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}
