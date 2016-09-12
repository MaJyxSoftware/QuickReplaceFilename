using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Quick_Replace_Filename
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog browse = new FolderBrowserDialog();
            if(browse.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = browse.SelectedPath;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(txtPath.Text) && !string.IsNullOrEmpty(txtSearch.Text))
            {
                string search = txtSearch.Text;
                string path = txtPath.Text;
                string replace = txtReplace.Text;
                rename(path, search, replace);
                MessageBox.Show("It's done!");
            }
        }

        private void rename(string path, string search, string replace)
        {
            if(search.Contains(";"))
            {
                string[] searchs = search.Split(';');
                string[] replaces = replace.Split(';');
                for(int i = 0; i < searchs.Length; i++)
                {
                    rename(path, searchs[i], replaces[i]);
                }
            } else
            {
                foreach (string file in Directory.GetFiles(path))
                {
                    if (Path.GetFileNameWithoutExtension(file).EndsWith(search))
                    {
                        //MessageBox.Show(Path.Combine(Path.GetDirectoryName(file), Path.GetFileName(file).Replace(search, replace)));
                        File.Move(file, Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file).Replace(search, replace) + Path.GetExtension(file)));
                    }
                }

                foreach (string dir in Directory.GetDirectories(path))
                {
                    rename(dir, search, replace);
                }
            }
        }

        private void btnBrowseE_Click(object sender, EventArgs e)
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.Filter = "Excel file (*.xlsx, *.xls)|*.xlsx;*.xls";

            if (browse.ShowDialog() == DialogResult.OK)
            {
                txtExcel.Text = browse.FileName;
                FileStream stream = File.Open(browse.FileName, FileMode.Open, FileAccess.Read);

                IExcelDataReader excelReader = null;
                if (Path.GetExtension(browse.FileName) == ".xls")
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                } else
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                excelReader.IsFirstRowAsColumnNames = true;
                DataSet result = excelReader.AsDataSet();

                List<string> search = new List<string>();
                List<string> replace = new List<string>();

                foreach (DataRow v in result.Tables[0].Rows)
                {
                    if (string.IsNullOrEmpty(v.ItemArray[0].ToString())) continue;

                    List<string> rTerm = new List<string>();
                    for(int i = 1; i < v.ItemArray.Length; i++)
                    {
                        rTerm.Add(v.ItemArray[i].ToString());
                    }

                    if (v.ItemArray[0].ToString().Contains("-"))
                    {
                        string[] sTerm = v.ItemArray[0].ToString().Split('-');
                        for (int i = 0; i < sTerm.Length; i++)
                        {
                            search.Add(sTerm[i]);
                            replace.Add(string.Join(" - ", rTerm) + " - " + (i + 1));
                        }
                    } else
                    {
                        search.Add(v.ItemArray[0].ToString());
                        replace.Add(string.Join(" - ", rTerm));
                    }
                }

                excelReader.Close();

                txtSearch.Text = string.Join(";", search);
                txtReplace.Text = string.Join(";", replace);

            }
        }
    }
}
