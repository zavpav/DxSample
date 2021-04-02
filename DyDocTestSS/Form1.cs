using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DyDocTestSS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.dynamicSheet1.TestLoadSheet();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            this.dynamicSheet1.ExportExcel("Экспорт.xlsx");
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            this.dynamicSheet1.ExportExcel("Экспорт.xlsx", true);
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
        }
    }
}
