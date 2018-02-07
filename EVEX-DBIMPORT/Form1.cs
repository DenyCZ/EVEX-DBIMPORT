using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace EVEX_DBIMPORT
{
    public partial class Form1 : Form
    {
        private string[] files;
        public General general;
        private List<ProjectList> ListPL;

        public Form1()
        {
            InitializeComponent();
            general = new General();
            ListPL = new List<ProjectList>();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                files = null;
                files = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.xlsx", SearchOption.AllDirectories);
            }

           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ListPL.Clear();

            foreach(string file in files)
            {
                if(file.Contains("_PL_") && !file.Contains("kopie"))
                {
                    ListPL.Add(new ProjectList(new ExcelPackage(new FileInfo(file))));
                }
            }

            ExcelUtils eUt = new ExcelUtils();

            foreach(ProjectList pl in ListPL)
            {
                ExcelRange range = pl.getPList().Workbook.Worksheets[1].Cells["B1"];
                //PLParser plParser = new PLParser(pl.getPList().Workbook);
            }

        }

    }
}
