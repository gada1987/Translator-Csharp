using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MyFirstWindos
{
    public partial class Form1 : Form
    {
        string defaultFilePath;
        string alternativeFilePath;
        string resultFilePath;
        string excelFilePath;
        string excelFilePath_com;
        string newAlternativeFilePath;
        bool button2WasClicked = true;
        bool button3WasClicked = true;
        bool button4WasClicked = true;
        bool button5WasClicked = true;
        bool button6WasClicked = true;
        bool button7WasClicked = true;
        bool button8WasClicked = true;
        string[] arraypath = new string[3];

        public Form1()
        { 

            InitAspose();
            InitializeComponent();

        }

        private void InitAspose()
        {
            try
            {
                var cells = new Aspose.Cells.License();
                cells.SetLicense("Aspose.Cells.lic");
            }
            catch (Exception e)
            {
                throw e;
            }
        }

/*===================================== Open Excel Sheet=======================================*/

        private void Button2_Click(object sender, EventArgs e)
        {
            if (button2WasClicked){
                dataGridView1.DataSource = null;
                 
                }

            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                openFileDialog1.Filter = "XML Files,Text Files, Excel Files|*.xlsx; *.xls; *.xml; *.txt;";

                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.tb_path.Text = openFileDialog1.FileName;
                    excelFilePath_com = this.tb_path.Text;
                    arraypath[2] = excelFilePath_com;

                }

                string constr = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + tb_path.Text + ";Extended Properties = \"Excel 12.0; HDR=Yes;\"; ");

                OleDbConnection con = new OleDbConnection(constr);
                con.Open();
                dropdown_sheet.DataSource = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                dropdown_sheet.DisplayMember = "TABLE_NAME";
                dropdown_sheet.ValueMember = "TABLE_NAME";
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message,
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }
            
            

        }
/*============================================ Laod_Sheet_1 =============================================*/
        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string constr = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + tb_path.Text + ";Extended Properties = \"Excel 12.0; HDR=Yes;\"; ");
                OleDbConnection con = new OleDbConnection(constr);
                OleDbDataAdapter sda = new OleDbDataAdapter("Select name, Value From[" + dropdown_sheet.SelectedValue + "]", con);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                foreach (DataRow row in dt.Rows)
                {
                    dataGridView1.DataSource = dt;
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
/*======================================== Open Default File =============================================*/
        private void Button3_Click(object sender, EventArgs e)
        {
            if (button3WasClicked)
            {
                dataGridView1.DataSource = null;
            }

            openFD.InitialDirectory = "C:";
            openFD.Title = "Insert an File";
            openFD.FileName = "DefaultFile";
            openFD.Filter = "XML Files |*.xml";
            openFD.FilterIndex = 1;
            openFD.RestoreDirectory = false;
            openFD.ReadOnlyChecked = true;
            openFD.ShowReadOnly = true;


            if (openFD.ShowDialog() == DialogResult.OK)
            {
               // Process.Start(openFD.FileName);
                DataSet ds = new DataSet();
                ds.ReadXml(openFD.FileName);
                dataGridView1.DataSource = ds.Tables[0];
                defaultFilePath = openFD.FileName;
                arraypath[0] = defaultFilePath;
            }

            else
            {

             MessageBox.Show("Operation Cancelled",
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }


        }
/*===================================== Open Alternative File ============================================*/
        private void Button4_Click(object sender, EventArgs e)
        {
            if (button4WasClicked == true)
            {
                dataGridView1.DataSource = null;
            }

            openFD.InitialDirectory = "C:";
            openFD.Title = "Insert an File";
            openFD.FileName = "AlternativeFile";
            openFD.Filter = "XML Files |*.xml";
            openFD.FilterIndex = 1;
            openFD.RestoreDirectory = true;
            openFD.ReadOnlyChecked = true;
            openFD.ShowReadOnly = true;


            if (openFD.ShowDialog() == DialogResult.OK)
            {
               // Process.Start(openFD.FileName);
                DataSet ds = new DataSet();
                ds.ReadXml(openFD.FileName);
                dataGridView1.DataSource = ds.Tables[0];
                alternativeFilePath = openFD.FileName;
                arraypath[1] = alternativeFilePath;
            }

            else
            {

             MessageBox.Show("Operation Cancelled",
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }


        }
        /*==============================================Show Result File =======================================*/
        private void Button5_Click(object sender, EventArgs e)
        {
            if (button5WasClicked) {
                dataGridView1.DataSource = null; }

            if (button6WasClicked && arraypath[0] != null && arraypath[1] != null)
            {
                DataSet ds = new DataSet();
                ds.ReadXml(resultFilePath);
                dataGridView1.DataSource = ds.Tables[0];
                //Process.Start(resultFilePath);
            }
            else {
                MessageBox.Show("Compare first to get a Result File!",
          "Important Note",
           MessageBoxButtons.OK,
           MessageBoxIcon.Exclamation,
           MessageBoxDefaultButton.Button1);
            }

            
        }

        /*========================================= Compare default_Alternative ================================*/
        private void Button6_Click(object sender, EventArgs e)
        {

          if (button6WasClicked) {
                dataGridView1.DataSource = null;
                dataGridView1.Refresh();
            }

            if (arraypath[0] == null)
            {
                MessageBox.Show("You need to Choose DefaultFilePath!",
                "Important Note",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1);
            }
            else if (arraypath[1] == null)
            {
             MessageBox.Show("You need to Choose AlternativeFilePath!",
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }


           else if (button6WasClicked && arraypath[0] != null && arraypath[1] != null){
              
               DialogResult messageresult = MessageBox.Show("Do you want to show Result File?", "Confirmation",MessageBoxButtons.YesNo);
                if (messageresult== DialogResult.Yes) {
                  
                    MessageBox.Show("Click ShowResult button please!");    
                }
                if (messageresult== DialogResult.No) {
                    MessageBox.Show("Welcome Back!");
                }
          

            resultFilePath = defaultFilePath.Substring(0, defaultFilePath.LastIndexOf("\\") + 1) + "ResultFile.xml";
             Compare_D_A d_A = new Compare_D_A();
             d_A.CompareD_A(defaultFilePath, alternativeFilePath, resultFilePath);
            }

        }
 /*===================================Compare default_Result====================================*/
        private void Button7_Click(object sender, EventArgs e)
        {
            if (button7WasClicked)
            {
                dataGridView1.DataSource = null;
                dataGridView1.Refresh();
            }
            if (arraypath[0] == null)
            {

                MessageBox.Show("You need to Choose DefaultFilePath!",
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }
            else if (arraypath[1] == null)
            {

                MessageBox.Show("You need to Choose AlternativeFilePath!",
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }
            else if (arraypath[2] == null)
            {

                MessageBox.Show("You need to Choose ExceltFilePath!",
             "Important Note",
             MessageBoxButtons.OK,
             MessageBoxIcon.Error,
             MessageBoxDefaultButton.Button1);
            }
            else if (button7WasClicked && arraypath[0] != null && arraypath[1] != null && arraypath[2] != null)
            {

                DialogResult messageresult = MessageBox.Show("Do you want to show A new Alternative File?", "Confirmation", MessageBoxButtons.YesNo);
                if (messageresult == DialogResult.Yes)
                {

                    MessageBox.Show("Click ShowAlternativeResult button please!");
                }
                if (messageresult == DialogResult.No)
                {
                    MessageBox.Show("Welcome Back!");
                }
               
                Compare_D_R d_R = new Compare_D_R();
                d_R.compareD_R(excelFilePath_com, defaultFilePath, alternativeFilePath);
                newAlternativeFilePath = alternativeFilePath.Substring(0, alternativeFilePath.LastIndexOf("\\") + 1) + "newAlternativeFile.xml";
                File.Copy(alternativeFilePath, newAlternativeFilePath, true);
            }
           

        }
        /*==========================================================================*/
        private void Button8_Click(object sender, EventArgs e)
        {
            if (button8WasClicked)
            {
                dataGridView1.DataSource = null;
                dataGridView1.Refresh();
            }

            if (button7WasClicked && arraypath[0] != null && arraypath[1] != null && arraypath[2] != null)
            {
                DataSet ds = new DataSet();
                ds.ReadXml(newAlternativeFilePath);
                dataGridView1.DataSource = ds.Tables[0];
                Process.Start(newAlternativeFilePath);
            }
            else
            {
                MessageBox.Show("Compare first to get a A new Alternative File!",
          "Important Note",
           MessageBoxButtons.OK,
           MessageBoxIcon.Exclamation,
           MessageBoxDefaultButton.Button1);
            }

        }



        /*================================================================================================*/


    }
}
