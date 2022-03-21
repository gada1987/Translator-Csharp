using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace TimeReport
{
    public partial class Form1 : Form
    {
        bool Sumbuttoncklicked = true;
        DataTable dataTable = new DataTable();
        DataTable dtGridSource = new DataTable();



        public Form1()
        {
            InitializeComponent();


        }
        class MyClass
        {
            public string TableName { get; set; }
            public string Path { get; set; }

        }

        //===============================Open===============================
        private void Button2_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable();

            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                openFileDialog1.Filter = "XML Files, Text Files, Excel Files| *.xlsx; *.xls; *.xml; *.txt; "; ;
                openFileDialog1.Multiselect = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    foreach (String file in openFileDialog1.FileNames)
                    {

                        tb_path.Text += file + ";";

                        string constr = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + file + ";Extended Properties = \"Excel 12.0; HDR=Yes;\"; ");

                        OleDbConnection con = new OleDbConnection(constr);
                        con.Open();

                        DataTable dt1 = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                        foreach (DataRow row in dt1.Rows)
                        {
                            string name = row["Table_Name"].ToString();
                            drop_down_sheet.Items.Add(new MyClass() { TableName = name, Path = file });
                        }
                        drop_down_sheet.DisplayMember = "TableName";

                    }

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }

        }
        //================================Import=====================
        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dataTable = new DataTable();
                System.Windows.Forms.ComboBox.ObjectCollection items = drop_down_sheet.Items;
                foreach (var item in items)
                {
                    MyClass myClass = (MyClass)item;

                    string constr = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + myClass.Path + ";Extended Properties = \"Excel 12.0; HDR=Yes;\"; ");
                    OleDbConnection con = new OleDbConnection(constr);
                    DataTable dt = new DataTable();
                    using (OleDbCommand comm = new OleDbCommand())
                    {

                        comm.CommandText = "Select * From[" + myClass.TableName + "]";
                        comm.Connection = con;
                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            da.SelectCommand = comm;
                            da.Fill(dt);
                            dataTable.Merge(dt);

                        }
                    }
                }
                dataGridView1.DataSource = dataTable;

            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);

            }
        }

        //======================Filter date to date=======================
        private void Button3_Click(object sender, EventArgs e)
        {

            try
            {
                DataTable dataTable = new DataTable();
                System.Windows.Forms.ComboBox.ObjectCollection items = drop_down_sheet.Items;
                foreach (var item in items)
                {

                    MyClass myClass = (MyClass)item;
                    //  MyClass myClass = (MyClass)drop_down_sheet.SelectedItem;
                    string constr = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + myClass.Path + ";Extended Properties = \"Excel 12.0; HDR=Yes;\"; ");
                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand command = new OleDbCommand("Select * From[" + myClass.TableName + "]", con);
                    OleDbDataAdapter da = new OleDbDataAdapter(command);

                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataTable.Merge(dt);


                }

                string filter = "Datum >= '" + startDate.Value.ToString("yyyy-MM-dd") + "' AND Datum <= '" + endDate.Value.ToString("yyyy-MM-dd") + "'";
                DataRow[] filteredRows = dataTable.Select(filter);
                dataGridView1.DataSource = filteredRows.CopyToDataTable();

            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);

            }


        }

//===============================Filter enligt Projekt, användare, Datum, aktivitet =============================

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

            if (comboBox1.Items[comboBox1.SelectedIndex].ToString() == "Projekt")
            {
                string filterField = "Projekt";
                DataTable dataTable = (DataTable)dataGridView1.DataSource;
                dataTable.DefaultView.RowFilter = string.Format("Convert([{0}], 'System.String') LIKE '%{1}%'", filterField, textBox2.Text);

            }

            if (comboBox1.Items[comboBox1.SelectedIndex].ToString() == "Användare")
            {

                string filter_An = "Användare";
                DataTable dataTable = (DataTable)dataGridView1.DataSource;
                dataTable.DefaultView.RowFilter = string.Format("Convert([{0}], 'System.String') LIKE '%{1}%'", filter_An, textBox2.Text);
            }

            if (comboBox1.Items[comboBox1.SelectedIndex].ToString() == "Aktivitet")
            {
                string filter_Ak = "Aktivitet";
                DataTable dataTable = (DataTable)dataGridView1.DataSource;
                dataTable.DefaultView.RowFilter = string.Format("Convert([{0}], 'System.String') LIKE '%{1}%'", filter_Ak, textBox2.Text);

            }

            if (comboBox1.Items[comboBox1.SelectedIndex].ToString() == "Datum")
            {

                string filterField = "Datum";
                DataTable dataTable = (DataTable)dataGridView1.DataSource;
                dataTable.DefaultView.RowFilter = string.Format("Convert([{0}], 'System.String') LIKE '%{1}%'", filterField, textBox2.Text);
            }
        }

 //====================== Sum =======================================
        private void Button3_Click_1(object sender, EventArgs e)

        {

            if (Sumbuttoncklicked)
            {
                decimal sum = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    sum += Convert.ToDecimal(dataGridView1.Rows[i].Cells["Timmar"].Value);
                }
                textBox1.Text = "Total Timmar =  " + sum.ToString();
            }
        }

  /*==============================Export =============================================*/
        private static void ExportToExcel(DataTable tbl, string excelFilePath, string starttime,string endtime)
        {
            try
            {
               
                if (tbl == null || tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var excelApp = new Excel.Application();
                excelApp.Workbooks.Add();

                // single worksheet
                Excel._Worksheet workSheet = excelApp.ActiveSheet;

                // column 
                for (var i = 0; i < tbl.Columns.Count; ++i)
                {
                    workSheet.Cells[2, i + 1] = tbl.Columns[i].ColumnName;
                }

                // rows
                for (var i = 1; i < tbl.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < tbl.Columns.Count; j++)
                    {
                        workSheet.Cells[i + 2, j + 1] = tbl.Rows[i][j];
                        if (j == tbl.Columns.Count - 1)
                        {
                            workSheet.Cells[1, 1] = "From: "+ starttime; 
                            workSheet.Cells[1, 2] = "To: " + endtime;
                          
                        }
                    }
                }

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        workSheet.SaveAs(excelFilePath);
                        excelApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                            + ex.Message);
                    }
                }
                else
                { // no file path is given
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

 /*==========================Projekt lista================================*/
        private void Button4_Click(object sender, EventArgs e)
        {

            DataTable dt = (DataTable)dataGridView1.DataSource;
            var data =
                from r in dt.AsEnumerable()
                group r by r.Field<string>("Projekt") into ProGroup

                let SumTime = ProGroup.Sum(r => Convert.ToDecimal (r.Field<string>("Timmar")))
                select new
                {
                    Projekt = ProGroup.Key,
                    Timmar = SumTime
                };
          
            var a = data.ToList();

            DataTable dt1 = new DataTable();
            dt1.Columns.Add("Projekt", typeof(string));
            dt1.Columns.Add("Timmar", typeof(decimal));
            foreach (var item in a)
            {
                dt1.Rows.Add(item.Projekt, item.Timmar);
            }

            dataGridView1.DataSource = dt1;
            ExportToExcel(dt1, @"C:\Users\gaad\Desktop\timeReport_App\TimeReport\Entergate_Data\Statistik_Projekt.xlsx", startDate.Value.ToString("yyyy-MM-dd") , endDate.Value.ToString("yyyy-MM-dd"));
           

            Console.WriteLine();
        }
 /*=====================Användare List =================================*/
        private void Button5_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            var data =
                from r in dt.AsEnumerable()
                group r by r.Field<string>("Användare") into ProGroup

                let SumTime = ProGroup.Sum(r => Convert.ToDecimal(r.Field<string>("Timmar")))
                select new
                {
                    Användare = ProGroup.Key,
                    Timmar = SumTime
                };
           
            var a = data.ToList();

            DataTable dt1 = new DataTable();
            dt1.Columns.Add("Användare", typeof(string));
            dt1.Columns.Add("Timmar", typeof(decimal));
            foreach (var item in a)
            {
                dt1.Rows.Add(item.Användare, item.Timmar);
            }

            dataGridView1.DataSource = dt1;
            ExportToExcel(dt1, @"C:\Users\gaad\Desktop\timeReport_App\TimeReport\Entergate_Data\Statistik_Användare.xlsx", startDate.Value.ToString("yyyy-MM-dd"), endDate.Value.ToString("yyyy-MM-dd"));
           

            Console.WriteLine();
        }


 //============================================================================================================================

        private void Button6_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            var data =
                from r in dt.AsEnumerable()
                group r by r.Field<string>("Aktivitet") into ProGroup

                let SumTime = ProGroup.Sum(r => Convert.ToDecimal(r.Field<string>("Timmar")))
                select new
                {
                    Aktivitet = ProGroup.Key,
                    Timmar = SumTime
                };

            var a = data.ToList();

            DataTable dt1 = new DataTable();
            dt1.Columns.Add("Aktivitet", typeof(string));
            dt1.Columns.Add("Timmar", typeof(decimal));
            foreach (var item in a)
            {
                dt1.Rows.Add(item.Aktivitet, item.Timmar);
            }

            dataGridView1.DataSource = dt1;
            ExportToExcel(dt1, @"C:\Users\gaad\Desktop\timeReport_App\TimeReport\Entergate_Data\Statistik_Aktivitet.xlsx", startDate.Value.ToString("yyyy-MM-dd"), endDate.Value.ToString("yyyy-MM-dd"));
         
            Console.WriteLine();
        }
    

        //======================================Un_used function====================================

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
         

        }
        
        private void Tb_path_TextChanged(object sender, EventArgs e)
        {

        }

        private void Drop_down_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }
        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void Label5_Click(object sender, EventArgs e)
        {

        }

       
    }

}




    
    

