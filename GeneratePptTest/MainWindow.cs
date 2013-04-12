using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GeneratePptTest.Business;

namespace GeneratePptTest
{
    public partial class MainWindow : Form
    {
        /// <summary>
        /// Gets or sets the template path.
        /// </summary>
        /// <value>
        /// The template path.
        /// </value>
        string TemplatePath { get; set; }

        /// <summary>
        /// Gets or sets the export PPT file path.
        /// </summary>
        /// <value>
        /// The export PPT file path.
        /// </value>
        string ExportPptFilePath { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            this.TemplatePath = @"Templates\template.ppt";
            this.ExportPptFilePath = @"D:\demoppt.ppt";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Form1Manager.GeneratePpt(this.TemplatePath, this.ExportPptFilePath))
            {
                MessageBox.Show("Template file filled with data was stored in " + this.ExportPptFilePath + "   PASSWORD is: asd");
            }
            else
            {
                MessageBox.Show("Template generation failed. Some exception may have occured. Please verify that the template and the data are correct");
            }
        }
    }
}
