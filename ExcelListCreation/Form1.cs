using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;





namespace ExcelListCreation
{
    public partial class UsersListCreation : Form
    {
        DateTime today = DateTime.Today;
        private string importedFilePath=null;
        private string exportFilePath=null;
        
        public UsersListCreation()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

        }

        //select initial list
        private void Button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog vi = new OpenFileDialog();
            vi.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (vi.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = vi.FileName;
                importedFilePath = vi.FileName;
                //Console.WriteLine(importedFilePath); 
                //RichTextBox.Text = Path.GetFileName(vi.FileName);

            }
        }
        //select file path
        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fld = new FolderBrowserDialog();
            if (fld.ShowDialog() == DialogResult.OK)
            {
                exportFilePath = (string)fld.SelectedPath;
                textBox2.Text = exportFilePath;
                //MessageBox.Show(exportFilePath);
            }
        }

        //creat the list
        private void Button2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (String.IsNullOrEmpty(textBox1.Text)) 
            {
                importedFilePath = null;
            }
            if (String.IsNullOrEmpty(textBox2.Text)) 
            {
                exportFilePath = null;
            }
            importedFilePath = textBox1.Text;
            exportFilePath = textBox2.Text;
             if (!String.IsNullOrEmpty(importedFilePath) && File.Exists(importedFilePath) && Directory.Exists(exportFilePath) && !String.IsNullOrEmpty(exportFilePath)) 
            {

                this.timer1.Start();
                List<RowOfImportedExcel> rowsOfImportedExcel = ImportInitialList.ReadFromExcel(importedFilePath);
                Console.WriteLine(rowsOfImportedExcel.Capacity);
                List<RowOfExportedExcel> rowsOfExportedExcel = CreationList.GenerateRowOfExportedExcel(rowsOfImportedExcel);
                ExportToExcelFile.ExportToExcel(rowsOfExportedExcel, exportFilePath);
                this.timer1.Stop();
                string filePath = exportFilePath + "\\" + ExportToExcelFile.fileName;
                string argument = "/select, \"" + filePath + "\"";
                Process.Start("explorer.exe",argument);
            }
            else 
            {
                if (!File.Exists(importedFilePath)&& !String.IsNullOrEmpty(importedFilePath)) 
                {
                    MessageBox.Show("File does not exist");
                }
                if (!Directory.Exists(exportFilePath)&& !String.IsNullOrEmpty(exportFilePath))
                {
                    MessageBox.Show("Folder does not exist");
                }
                if (String.IsNullOrEmpty(importedFilePath))
                {
                    MessageBox.Show("Please choose an excel file");
                }
                else if (String.IsNullOrEmpty(exportFilePath)) 
                {
                    MessageBox.Show("Please choose a folder");
                }
           
            }
            Cursor.Current = Cursors.Default;

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Increment(1);
        }
       
        private void Help_Click(object sender, EventArgs e)
        {
            {
                Form form2 = new Form();
                form2.Show();
                form2.Text = "Initial list example";

                string pictureFile = Path.GetDirectoryName(Application.ExecutablePath)+ "\\correctFormForUsersListCreation.bmp";
                Console.WriteLine(pictureFile);
                PictureBox PictureBox1 = new PictureBox();
                PictureBox1.Image = new Bitmap(pictureFile);
                //PictureBox1.Size = form2.Size;
                PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
                PictureBox1.Location = new Point(0, 12);
                PictureBox1.TabIndex = 7;
                form2.Width = PictureBox1.Width;
                // Add the new control to its parent's controls collection
               
                
                form2.Controls.Add(PictureBox1);
                

               

            }
            
        }

        

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
