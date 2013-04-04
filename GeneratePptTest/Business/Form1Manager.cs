using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointGenerator.Helpers;
using PowerPointGenerator.Managers;

namespace GeneratePptTest.Business
{
    public static class Form1Manager
    {
        public static bool GeneratePpt(string templatePath, string exportPptFilePath)
        {
            string connectionString = ConfigHelper.GenerateConnectionStringMssql();
            StoreManager store = new StoreManager(connectionString);
            return store.GeneratePpt(templatePath, exportPptFilePath);
        }
    }
}
