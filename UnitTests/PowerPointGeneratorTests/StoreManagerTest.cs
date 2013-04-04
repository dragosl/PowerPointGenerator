using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using PowerPointGenerator.Helpers;
using PowerPointGenerator.Managers;
using PowerPointGenerator.Model;

namespace UnitTests.PowerPointGeneratorTests
{
    [TestFixture]
    public class StoreManagerTest
    {
        /// <summary>
        /// The DB valid connection.
        /// </summary>
        SqlConnection connection;

        /// <summary>
        /// The DB invalid connection.
        /// </summary>
        SqlConnection invalidConnection;

        /// <summary>
        /// The sales number.
        /// </summary>
        int salesNumber;

        /// <summary>
        /// The first sale.
        /// </summary>
        Sale firstSale;

        /// <summary>
        /// The connection string
        /// </summary>
        string connectionString;

        /// <summary>
        /// The constructor sales
        /// </summary>
        List<Sale> constructorSales;

        /// <summary>
        /// The invalid connection string
        /// </summary>
        string invalidConnectionString;

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

        [SetUp]
        public void Init()
        {
            connectionString = ConfigHelper.GenerateConnectionStringMssql();
            connection = new SqlConnection(connectionString);

            invalidConnectionString = "server=DRAGOSL-PC\\SQLEXS;database=AdventureWorks2008;";

            salesNumber = 701;

            // first sale from DB
            firstSale = new Sale
            {
                BusinessEntityID = 292,
                Demographics = "<StoreSurvey xmlns=\"http://schemas.microsoft.com/sqlserver/2004/07/adventure-works/StoreSurvey\"><AnnualSales>800000</AnnualSales><AnnualRevenue>80000</AnnualRevenue><BankName>United Security</BankName><BusinessType>BM</BusinessType><YearOpened>1996</YearOpened><Specialty>Mountain</Specialty><SquareFeet>21000</SquareFeet><Brands>2</Brands><Internet>ISDN</Internet><NumberEmployees>13</NumberEmployees></StoreSurvey>",
                ModifiedDate = new DateTime(2004, 10, 13, 11, 15, 07),
                Name = "Next-Door Bike Store",
                Rowguid = "a22517e3-848d-4ebe-b9d9-7437f3432304",
                SalesPersonID = 279
            };

            constructorSales = new List<Sale> { firstSale };

            templatePath = @"Templates\template.ppt";
            savePptFilePath = @"demoppt.ppt";

            invalidTemplatePath = @"Templates\template.pptx";
        }

        /// <summary>
        /// Store constructor sales property count test.
        /// </summary>
        [Test]
        public void StoreConstructorSalesPropertyCountTest()
        {
            StoreManager store = new StoreManager(connectionString);
            Assert.AreEqual(store.Sales.Count, salesNumber);
        }

        /// <summary>
        /// Store constructor sales property first business entity ID equal test.
        /// </summary>
        [Test]
        public void StoreConstructorSalesPropertyFirstBusinessEntityIDEqualTest()
        {
            StoreManager store = new StoreManager(connectionString);
            Assert.AreEqual(store.Sales[0].BusinessEntityID, firstSale.BusinessEntityID);
        }

        /// <summary>
        /// Store constructor sales property first sales person ID equal test.
        /// </summary>
        [Test]
        public void StoreConstructorSalesPropertyFirstSalesPersonIDEqualTest()
        {
            StoreManager store = new StoreManager(connectionString);
            Assert.AreEqual(store.Sales[0].SalesPersonID, firstSale.SalesPersonID);
        }

        /// <summary>
        /// Store constructor sales property exception test.
        /// </summary>
        [Test]
        [ExpectedException(typeof(PowerPointGenerator.Exceptions.SqlException))]
        public void StoreConstructorSalesPropertyExceptionTest()
        {
            StoreManager store = new StoreManager(invalidConnectionString);
        }

        /// <summary>
        /// Store constructor with data equal sales test.
        /// </summary>
        [Test]
        public void StoreConstructorWithDataEqualSalesTest()
        {
            StoreManager store = new StoreManager(constructorSales);
            Assert.AreEqual(store.Sales, constructorSales);
        }

        /// <summary>
        /// Test which verifies if the ppt was generated with success.
        /// </summary>
        [Test]
        public void GeneratePptTest()
        {
            StoreManager store = new StoreManager(connectionString);
            Assert.IsTrue(store.GeneratePpt(templatePath, savePptFilePath));
        }

        /// <summary>
        /// Test which verifies if the ppt failed to be generated, because of the template inconsistency.
        /// </summary>
        [Test]
        public void GeneratePptInvalidTemplateTest()
        {
            StoreManager store = new StoreManager(connectionString);
            Assert.IsFalse(store.GeneratePpt(invalidTemplatePath, savePptFilePath));
        }
    }
}
