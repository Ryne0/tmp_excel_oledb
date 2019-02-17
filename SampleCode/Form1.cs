using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Excel_using_OleDB
{
    public partial class Form1 : Form
    {

        string connectionString = "";
        OleDbCommand cmd = null;
        OleDbConnection objCon = null;
        OleDbDataAdapter objDA = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void cmdBrowse_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            openFileDialog1.Filter = "Excel files|*.xls;*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                txtFilePath.Text = openFileDialog1.SafeFileName;
            }
            Console.WriteLine(result);
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            Environment.Exit(1);
        }

        private void cmdReadExcel_Click(object sender, EventArgs e)
        {

            DataSet objDS = new DataSet();

            string szConStr = collect_Con(System.IO.Path.GetExtension(txtFilePath.Text), txtFilePath.Text);


            try
            {
                objCon = new OleDbConnection(szConStr);
                objCon.Open();
                objDA = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", objCon);
                DataTable objDTExcel = new DataTable();
                objDA.Fill(objDTExcel);
                dataGridView1.DataSource = objDTExcel;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (objDA != null)
                {
                    objDA.Dispose();
                    objDA = null;
                }

                if (objCon != null)
                {
                    objCon.Close();
                    objCon.Dispose();
                    objCon = null;
                }
            }


        }



        private string collect_Con(string szFileExtension, string szPath)
        {

            switch (szFileExtension.ToUpper())
            {
                case ".XLS":
                    connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=
                                        " + szPath + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                    break;
                case ".XLSX":
                    connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=
                                        " + szPath + ";Extended Properties=\"Excel 12.0;HDR=YES;\"";
                    break;
                default:
                    break;
            }

            return connectionString;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtFilePath.Text = Application.StartupPath + "\\TestBook.xlsx";
            connectionString = collect_Con(System.IO.Path.GetExtension(txtFilePath.Text), txtFilePath.Text);
        }

        private void cmdWriteInExcel_Click(object sender, EventArgs e)
        {
            try
            {
                objCon = new OleDbConnection(connectionString);
                objCon.Open();
                cmd = new OleDbCommand();
                cmd.Connection = objCon;

                cmd.CommandText = "Insert into [Sheet1$] (month,mango,apple,orange) VALUES('DEC','40','60','80');";
                cmd.ExecuteNonQuery();

            }
            catch (Exception)
            {
                //exception here
            }
            finally
            {
                objCon.Close();
                objCon.Dispose();
            }
        }

        private void cmdUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                objCon = new OleDbConnection(connectionString);
                objCon.Open();
                cmd = new OleDbCommand();
                cmd.Connection = objCon;
                cmd.CommandText = "UPDATE [Sheet1$] SET month = 'DEC' WHERE apple = 60;";
                cmd.ExecuteNonQuery();

            }
            catch (Exception)
            {
                //exception here
            }
            finally
            {
                objCon.Close();
                objCon.Dispose();
            }

        }
    }
}
