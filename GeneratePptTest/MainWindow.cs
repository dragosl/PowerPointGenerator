using System;
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
            if (MainWindowManager.GeneratePpt(this.TemplatePath, this.ExportPptFilePath))
            {
                MessageBox.Show(string.Format(Properties.Resources.TemplateGenerateOkMessage, this.ExportPptFilePath));
            }
            else
            {
                MessageBox.Show(Properties.Resources.TemplateGenerateFailMessage);
            }
        }
    }
}
