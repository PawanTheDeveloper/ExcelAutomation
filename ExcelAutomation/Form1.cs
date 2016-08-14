using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel=Microsoft.Office.Interop.Excel;


namespace ExcelAutomation
{
    public partial class Form1 : Form
    {
        #region Data Members
        private List<int> columnIndex = new List<int>();

        private ExcelOperation objectExcelOperation = new ExcelOperation();

        #endregion


        public Form1()
        {
            InitializeComponent();
            initialize();            
        }

        private void initialize()
        {
            //button_add.Visible = false;
            //button_remove.Visible = false;
            //listBox_populatecolumn.Visible = false;
            //listBox_selectedcolumn.Visible = false;
            //button_generatefile.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = string.Empty;
            OpenFileDialog file_dialog = new OpenFileDialog();
            file_dialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            
            DialogResult result = file_dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox_firstfilename.Text = file_dialog.FileName;
                fileName = textBox_firstfilename.Text;
            }
            //string newFileName = fileName.Substring((5));
            //newFileName = newFileName + "temp.xlsb";
            string newFileName = @"C:\temp.xlsx";
            objectExcelOperation.Initialize(fileName,newFileName);
            populateColumns();
            file_dialog.Dispose();
        }
       
        private void populateColumns()
        {            
            List<string> columnPopulated=new List<string>(); 
            listBox_populatecolumn.DataSource = objectExcelOperation.ReadFirstRow(columnPopulated); ;
            //button_add.Visible = true;
            //button_generatefile.Visible = true;
            //listBox_selectedcolumn.Visible = true;
            //button_remove.Visible = true;
            //listBox_populatecolumn.Visible = true;
        }
                       
        private void button_add_Click(object sender, EventArgs e)
        {
            listBox_selectedcolumn.Items.Add(listBox_populatecolumn.SelectedItem);            
        }

        private void button_remove_Click(object sender, EventArgs e)
        {
            if (listBox_selectedcolumn.SelectedItem != null)
            {
                listBox_selectedcolumn.Items.Remove(listBox_selectedcolumn.SelectedItem);
            }
            else
            {
                MessageBox.Show("Please select an item to be deleted");
            }
        }

        private void getOrderOfColumnsInsideColumnSelectedList()
        {
            foreach (object item in listBox_selectedcolumn.Items)
            {
                int index = listBox_populatecolumn.Items.IndexOf(item);
                columnIndex.Add(index);
            }
        }
        private void button_generatefile_Click(object sender, EventArgs e)
        {
            bool answer;
            getOrderOfColumnsInsideColumnSelectedList();
            objectExcelOperation.setColumnUsedForComparison(columnIndex);
            int maxColumn=objectExcelOperation.getMaxRow();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = maxColumn;
            for (int row = 1; row <= maxColumn; row++)
            {
                answer=objectExcelOperation.Operation(row);
                progressBar1.Value++;
            }
            objectExcelOperation.WriteTheDate();
            objectExcelOperation.Quit();
            MessageBox.Show("File Generation Successful","Success",MessageBoxButtons.OK,MessageBoxIcon.Information);
            Application.Exit();
        }
    }
}
