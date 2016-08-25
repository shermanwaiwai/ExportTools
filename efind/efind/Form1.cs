using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Net;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
namespace efind
{
    public partial class Form1 : Form
    {
        public static string Path = "c:\\efind\\";
        public static string downloadlink = "http://13.186.65.9/RECTIDOC/docview_save.asp?DOWNLOAD=Y&DOC_ID=";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            //try
            //{
            //    Excel excel_ = new Excel();
            //    //excel_.AddHeaderToExcel();
            //    excel_.Create_Excel();
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
            //finally
            //{
            //    richTextBox1.AppendText("Export To Excel");
            //}
        }



        private void button4_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            SqlConnection connection;
            //SqlCommand command;
            //string SQL = "select * from obj as a left join (select a.obj_id , a.parent_id , obj_name , obj_create_time, obj_modify_time , obj_owner ,temp_id , doc_ext, doc_type, doc_version , doc_extpath from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D'" ;
            //SQL = SQL + ") as b on  a.parent_id = b.temp_id where a.parent_id in (select temp_id from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D')";
            
            //get the root folder
            string SQL = "select * from obj where parent_id = 0 and mark_del ='N' and obj_type = 'F' order by obj_id ";

            string datasource = textBox1.Text;
            string database = textBox2.Text;
            string userid = textBox3.Text;
            string password = textBox4.Text;

            //string password = "x-admin";
            connetionString = "Data Source=" + datasource +";Initial Catalog=" +database +";User ID=" +userid +";Password=" + password;
            connection = new SqlConnection(connetionString);
            
            SqlDataAdapter dscmd = new SqlDataAdapter(SQL, connection);
            DataSet ds = new DataSet();
            dscmd.Fill(ds);

            //string temp_SQL1_1 = "select a.obj_id as folderID, a.temp_id  , b.* from folder as a left join obj as b on  a.obj_id =  b.obj_id";
            //select folder sql
            string temp_SQL1_1 = "select q.obj_name , k.* from obj as q left join (select a.obj_id as folderID, a.temp_id  , b.* from folder as a left join obj as b on  a.obj_id =  b.obj_id where mark_del = 'N') as k on q.obj_id = k.temp_id where k.obj_type ='F' and q.mark_del ='N' ";
            SqlDataAdapter ada_1 = new SqlDataAdapter(temp_SQL1_1, connection);
            DataSet dataset_for_folder = new DataSet();
            ada_1.Fill(dataset_for_folder);

            int Current_Root_objID;
            string FullPathName = "";

            //string SQL_for_index = "select a.obj_id as index_parent_id , a.obj_name as index_name , b.* from obj as a left join (select a.obj_id as docID, a.parent_id , obj_name , obj_create_time, obj_modify_time , obj_owner ,temp_id , doc_ext, doc_type, doc_version , doc_extpath from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D') as b on  a.parent_id = b.temp_id where a.parent_id in (select temp_id from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D') order by docID";

            //string SQL_for_index = "select a.obj_id as index_parent_id , a.obj_name as index_name , b.* from obj as a left join (select a.obj_id as docID, a.parent_id , obj_name , obj_create_time, obj_modify_time , obj_owner ,temp_id , doc_ext, doc_type, doc_version , doc_extpath from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D') as b on  a.parent_id = b.temp_id where a.parent_id in (select temp_id from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D') order by docID";

            string SQL_for_index ="select a.obj_id as index_parent_id , a.obj_name as index_name , b.* from obj as a left join (select a.obj_id as docID, a.parent_id , obj_name , obj_create_time, obj_modify_time , obj_owner ,temp_id , doc_ext, doc_type, doc_version , doc_extpath from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D') as b on  a.parent_id = b.temp_id where a.parent_id in (select temp_id from obj as a left join document as b on a.obj_id = b.obj_id where a.obj_type ='D' and a.mark_del ='N' )  order by docID";

            SqlDataAdapter cmd = new SqlDataAdapter(SQL_for_index, connection);
            DataSet datasetfor_index = new DataSet();
            cmd.Fill(datasetfor_index);


            //string SQL_for_index_doc = "select  q.obj_id as docID, q.temp_id , q.field_id , q.field_valutxt, k.obj_name from (select a.obj_id , a.temp_id, a.field_id , a.field_valutxt  from content as a left join obj as b on a.obj_id = b.obj_id ) as q left join obj as k on q.field_id= k.obj_id and k.mark_del='N' order by docID";
            string SQL_for_index_doc = "select  q.obj_id as docID, q.temp_id , q.field_id , q.field_valutxt, k.obj_name from (select a.obj_id , a.temp_id, a.field_id , a.field_valutxt  from content as a left join obj as b on a.obj_id = b.obj_id where b.mark_del ='N' ) as q left join obj as k on q.field_id= k.obj_id  order by docID";

            SqlDataAdapter cmd_ = new SqlDataAdapter(SQL_for_index_doc, connection);
            DataSet dataset_for_index_doc = new DataSet();
            cmd_.Fill(dataset_for_index_doc);

            
            CreateFolder(Path);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                //get the obj 
                Current_Root_objID = Int32.Parse(ds.Tables[0].Rows[i].ItemArray[0].ToString());
                FullPathName = ds.Tables[0].Rows[i].ItemArray[3].ToString();
                FullPathName = FullPathName + "_" + Current_Root_objID;
                FullPathName = string.Format("{0}{1}", Path, FullPathName);
                string type = ds.Tables[0].Rows[i].ItemArray[2].ToString();
                if (type == "F")
                {
                    CreateFolder(FullPathName);
                }
                
                
                string temp_SQL = "select * from obj where mark_del ='N' and  parent_id = " + Current_Root_objID;
                SqlDataAdapter ada = new SqlDataAdapter(temp_SQL, connection);
                DataSet ds1 = new DataSet();
                ada.Fill(ds1);

                RecursiveFunction(ds1, datasetfor_index, dataset_for_folder, connection, FullPathName, dataset_for_index_doc);
               
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "RDTEST1\\RECTIDOC";
            textBox2.Text = "RECTIDOC";
            textBox3.Text = "sa";
            textBox4.Text = "x-admin";
        }

        public void WriteToRichBox(string message)
        {
            richTextBox1.AppendText(message);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            WebClient webClient = new WebClient();
            webClient.DownloadFile("http://13.186.65.9/RECTIDOC/docview_save.asp?DOWNLOAD=Y&DOC_ID=3905", @"c:\\3905_.pdf");
        }

        public bool DownloadFile(string DocID, string DocName, string FullPath)
        {
            try
            {
                WebClient webClient = new WebClient();
                string temp_downloadlink = "http://13.186.65.9/RECTIDOC/docview_save.asp?DOWNLOAD=Y&DOC_ID=" + DocID;
                //string temp_fileName = string.Format("{0}\\{1}", FullPath ,DocID);
                webClient.DownloadFile(temp_downloadlink, @FullPath);
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private void CreateFolder(string path)
        {
            bool folderExists = Directory.Exists(path);
            if (!folderExists)
                Directory.CreateDirectory(path);
        }

        public void CreateXml(DataSet dataset_for_index_doc, DataSet dataset_for_folder,  string indexing, string fullPath ,string filename, SqlConnection connection ,string folderID , string docID)
        {
            string temp_string = System.IO.Path.GetExtension(filename);
            string output_filename = filename.Replace(temp_string, ".xml");

            string temp_fullpath = string.Format("{0}\\{1}", fullPath, output_filename);

            //String SQLCommand = " select * from obj where parent_id =" + indexing ;

            DataRow[] indexRow = dataset_for_index_doc.Tables[0].Select("DocID = " + docID);

            XmlDataDocument doc = new XmlDataDocument();
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);
            DataRow[] r = dataset_for_folder.Tables[0].Select("folderID = " + folderID);
            string temp_index_name = r[0].ItemArray[0].ToString().Replace(" ","");
            XmlElement element1 = doc.CreateElement(string.Empty, temp_index_name, string.Empty);
            doc.AppendChild(element1);
            //Create an element representing the first customer record.
            //foreach (DataRow row in ds.Tables[0].Rows)
            //{
            //    string temp_index_name1 = row.ItemArray[3].ToString().Replace(" ", "");
            //    XmlElement element2 = doc.CreateElement(string.Empty, temp_index_name1, string.Empty);
            //    XmlText text1 = doc.CreateTextNode(row.ItemArray[3].ToString());
            //    element2.AppendChild(text1);
            //    element1.AppendChild(element2);
            //}
            foreach(DataRow index in indexRow )
            {
                string temp_index_name1 = index.ItemArray[4].ToString().Replace(" ", "");
                XmlElement element2 = doc.CreateElement(string.Empty, temp_index_name1, string.Empty);
                XmlText text1 = doc.CreateTextNode(index.ItemArray[3].ToString());
                element2.AppendChild(text1);
                element1.AppendChild(element2);
            }
            doc.Save(temp_fullpath);
            
        }

        public void CreateXml_old(DataRow[] dtRow, string fullPath)
        {
            XmlDocument doc = new XmlDocument();
            string[] index = new string[20];
            string temp_fullpath = string.Format("{0}\\{1}", fullPath, dtRow[0].ItemArray[4].ToString() + ".xml");
            //(1) the xml declaration is recommended, but not mandatory
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //(2) string.Empty makes cleaner code
            XmlElement element1 = doc.CreateElement(string.Empty, "body", string.Empty);
            doc.AppendChild(element1);

            XmlElement element2 = doc.CreateElement(string.Empty, "Index", string.Empty);
            element1.AppendChild(element2);

            foreach (DataRow row in dtRow)
            {
                int count = 0;
                index[count] = row.ItemArray[1].ToString();
                count++;
            }

            XmlElement element3 = doc.CreateElement(string.Empty, "index1", string.Empty);
            XmlText text1 = doc.CreateTextNode(index[0]);
            element3.AppendChild(text1);
            element2.AppendChild(element3);

            XmlElement element4 = doc.CreateElement(string.Empty, "index2", string.Empty);
            XmlText text2 = doc.CreateTextNode(index[1]);
            element4.AppendChild(text2);
            element2.AppendChild(element4);

            XmlElement element5 = doc.CreateElement(string.Empty, "index3", string.Empty);
            XmlText text3 = doc.CreateTextNode(index[2]);
            element5.AppendChild(text3);
            element2.AppendChild(element5);

            XmlElement element6 = doc.CreateElement(string.Empty, "index4", string.Empty);
            XmlText text4 = doc.CreateTextNode(index[3]);
            element6.AppendChild(text4);
            element2.AppendChild(element6);

            XmlElement element7 = doc.CreateElement(string.Empty, "index5", string.Empty);
            XmlText text5 = doc.CreateTextNode(index[4]);
            element7.AppendChild(text5);
            element2.AppendChild(element7);

            XmlElement element8 = doc.CreateElement(string.Empty, "index6", string.Empty);
            XmlText text6 = doc.CreateTextNode(index[5]);
            element8.AppendChild(text6);
            element2.AppendChild(element8);

            XmlElement element9 = doc.CreateElement(string.Empty, "index7", string.Empty);
            XmlText text7 = doc.CreateTextNode(index[6]);
            element9.AppendChild(text7);
            element2.AppendChild(element9);

            XmlElement element10 = doc.CreateElement(string.Empty, "index8", string.Empty);
            XmlText text8 = doc.CreateTextNode(index[7]);
            element10.AppendChild(text8);
            element2.AppendChild(element10);

            XmlElement element11 = doc.CreateElement(string.Empty, "index9", string.Empty);
            XmlText text9 = doc.CreateTextNode(index[8]);
            element11.AppendChild(text9);
            element2.AppendChild(element11);

            XmlElement element12 = doc.CreateElement(string.Empty, "index10", string.Empty);
            XmlText text10 = doc.CreateTextNode(index[9]);
            element12.AppendChild(text10);
            element2.AppendChild(element12);

            XmlElement element13 = doc.CreateElement(string.Empty, "index11", string.Empty);
            XmlText text11 = doc.CreateTextNode(index[10]);
            element13.AppendChild(text11);
            element2.AppendChild(element13);

            XmlElement element14 = doc.CreateElement(string.Empty, "index12", string.Empty);
            XmlText text12 = doc.CreateTextNode(index[11]);
            element14.AppendChild(text12);
            element2.AppendChild(element14);

            XmlElement element15 = doc.CreateElement(string.Empty, "index13", string.Empty);
            XmlText text13 = doc.CreateTextNode(index[12]);
            element15.AppendChild(text13);
            element2.AppendChild(element15);

            XmlElement element16 = doc.CreateElement(string.Empty, "index14", string.Empty);
            XmlText text14 = doc.CreateTextNode(index[13]);
            element16.AppendChild(text14);
            element2.AppendChild(element16);

            XmlElement element17 = doc.CreateElement(string.Empty, "index15", string.Empty);
            XmlText text15 = doc.CreateTextNode(index[14]);
            element17.AppendChild(text15);
            element2.AppendChild(element17);

            XmlElement element18 = doc.CreateElement(string.Empty, "index16", string.Empty);
            XmlText text16 = doc.CreateTextNode(index[15]);
            element18.AppendChild(text16);
            element2.AppendChild(element18);

            XmlElement element19 = doc.CreateElement(string.Empty, "index17", string.Empty);
            XmlText text17 = doc.CreateTextNode(index[16]);
            element19.AppendChild(text17);
            element2.AppendChild(element19);

            XmlElement element20 = doc.CreateElement(string.Empty, "index18", string.Empty);
            XmlText text18 = doc.CreateTextNode(index[17]);
            element20.AppendChild(text18);
            element2.AppendChild(element20);

            XmlElement element21 = doc.CreateElement(string.Empty, "index19", string.Empty);
            XmlText text19 = doc.CreateTextNode(index[18]);
            element21.AppendChild(text19);
            element2.AppendChild(element21);

            XmlElement element22 = doc.CreateElement(string.Empty, "index20", string.Empty);
            XmlText text20 = doc.CreateTextNode(index[19]);
            element22.AppendChild(text20);
            element2.AppendChild(element22);

            doc.Save(temp_fullpath);
        }

        public void RecursiveFunction(DataSet ds, DataSet datasetfor_index, DataSet dataset_for_folder, SqlConnection connection, string FullPathName, DataSet dataset_for_index_doc)
        {
            for (int ab = 0; ab < ds.Tables[0].Rows.Count; ab++)
            {
                string temp_type = ds.Tables[0].Rows[ab].ItemArray[2].ToString();
                string current_objIDj = ds.Tables[0].Rows[ab].ItemArray[0].ToString();
                string filename = ds.Tables[0].Rows[ab].ItemArray[3].ToString();
                string folderID = ds.Tables[0].Rows[ab].ItemArray[1].ToString();

                if (temp_type == "F")
                {
                    filename = filename + "_" + current_objIDj;
                }
                string temp_fullpath = string.Format("{0}\\{1}", FullPathName, filename);
                if (temp_type == "F")
                {
                    CreateFolder(temp_fullpath);
                }
                else if (temp_type == "D")
                {
                    //string folderID = ds.Tables[0].Rows[ab].ItemArray[1].ToString();
                    DataRow[] row = dataset_for_folder.Tables[0].Select("folderID = " + folderID);
                    string indexing =  row[0].ItemArray[2].ToString();

                    //var foundRows = dataset_for_folder.Tables[0].Select("obj_id = " + current_objIDj);
                    
                    //Create an element representing the first customer record.

                    //XmlElement elem = doc.GetElementFromRow(row);
                    //Console.WriteLine(elem.OuterXml);


                   

                    CreateXml(dataset_for_index_doc, dataset_for_folder, indexing, FullPathName, filename, connection, folderID, current_objIDj);
                    DownloadFile(current_objIDj, filename, temp_fullpath);
                }

                string temp_SQL1 = "select * from obj where  mark_del = 'N' and parent_id = " + current_objIDj ;
                SqlDataAdapter ada1 = new SqlDataAdapter(temp_SQL1, connection);
                DataSet ds2 = new DataSet();
                ada1.Fill(ds2);

                RecursiveFunction(ds2, datasetfor_index, dataset_for_folder, connection, temp_fullpath, dataset_for_index_doc);
            }
            return;
        }
    }
}
                

