using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPointGenerator.Helpers;
using PowerPointGenerator.Managers;

namespace GeneratePptTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = ConfigHelper.GenerateConnectionString();
            StoreManager store = new StoreManager(connectionString);
            store.GeneratePpt(@"Templates\template.ppt", @"demoppt.ppt");
            MessageBox.Show("Template file filled with data was stored as demoppt.ppt in the bin directory...");
        }
    }
}
