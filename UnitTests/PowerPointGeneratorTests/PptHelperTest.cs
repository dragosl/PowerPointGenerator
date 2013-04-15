using System.Collections.Generic;
using System.Data.SqlClient;
using GeneratePptTest.Helpers;
using NUnit.Framework;
using PowerPointGenerator.Helpers;
using PowerPointGenerator.Model;

namespace UnitTests.PowerPointGeneratorTests
{
    [TestFixture]
    public class PptHelperTest
    {
        /// <summary>
        /// The sales
        /// </summary>
        List<Sale> sales;

        /// <summary>
        /// The template path
        /// </summary>
        string templatePath;

        /// <summary>
        /// The save PPT file path
        /// </summary>
        string savePptFilePath;

        /// <summary>
        /// The invalid template path
        /// </summary>
        string invalidTemplatePath;

        /// <summary>
        /// Inits this instance.
        /// </summary>
        [SetUp]
        public void Init()
        {
            string connectionString = ConfigHelper.GenerateConnectionStringMssql();
            SqlConnection connection = new SqlConnection(connectionString);

            sales = SqlHelper.GetSales(connection);

            templatePath = @"Templates\template.ppt";
            savePptFilePath = @"demoppt.ppt";

            invalidTemplatePath = @"Templates\template.pptx";
        }

        /// <summary>
        /// Inserts the sales in template test.
        /// </summary>
        [Test]
        public void InsertSalesInTemplateTest()
        {
            Assert.AreEqual(PptHelper.InsertSalesInTemplate(sales, templatePath, savePptFilePath), true);
        }

        /// <summary>
        /// Inserts the sales in template fail test.
        /// </summary>
        [Test]
        public void InsertSalesInTemplateFailTest()
        {
            Assert.AreEqual(PptHelper.InsertSalesInTemplate(sales, invalidTemplatePath, savePptFilePath), false);
        }
    }
}
