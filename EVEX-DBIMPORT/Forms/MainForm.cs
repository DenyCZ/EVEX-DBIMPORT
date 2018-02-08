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
        private List<ProjectListWS> ListPL;

        public Form1()
        {
            InitializeComponent();
            general = new General();
            ListPL = new List<ProjectListWS>();

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
            files = null;
            files = Directory.GetFiles("D:\\ITIM\\EVEX", "*.xlsx", SearchOption.AllDirectories);

            ListPL.Clear();

            foreach (string file in files)
            {
                if (file.Contains("_PL_") && !file.Contains("kopie"))
                {
                    ListPL.Add(new ProjectListWS(new ExcelPackage(new FileInfo(file))));
                }
            }

            ExcelUtils eUt = new ExcelUtils();

            foreach (ProjectListWS pl in ListPL)
            {
                PLParser plParser = new PLParser(pl.getPList().Workbook);
            }

        }

    }
}
