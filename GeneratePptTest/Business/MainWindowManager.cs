using GeneratePptTest.Helpers;
using PowerPointGenerator.Managers;

namespace GeneratePptTest.Business
{
    public static class MainWindowManager
    {
        public static bool GeneratePpt(string templatePath, string exportPptFilePath)
        {
            string connectionString = ConfigHelper.GenerateConnectionStringMssql();
            StoreManager store = new StoreManager(connectionString);
            return store.GeneratePpt(templatePath, exportPptFilePath);
        }
    }
}
