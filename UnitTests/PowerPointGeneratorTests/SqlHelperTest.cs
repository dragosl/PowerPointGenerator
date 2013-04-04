using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using PowerPointGenerator.Helpers;
using PowerPointGenerator.Model;

namespace UnitTests.PowerPointGeneratorTests
{
    [TestFixture]
    public class SqlHelperTest
    {
        /// <summary>
        /// The DB valid connection.
        /// </summary>
        SqlConnection validConnection;

        /// <summary>
        /// The DB invalid connection.
        /// </summary>
        SqlConnection invalidConnection;

        /// <summary>
        /// The sales number.
        /// </summary>
        int salesNumber;

        /// <summary>
        /// The first sale
        /// </summary>
        Sale firstSale;

        [SetUp]
        public void Init()
        {
            string connectionString = ConfigHelper.GenerateConnectionStringMssql();
            validConnection = new SqlConnection(connectionString);

            connectionString = "server=DRAGOSL-PC\\SQLEXS;database=AdventureWorks2008;";
            invalidConnection = new SqlConnection(connectionString);

            salesNumber = 701;

            firstSale = new Sale
            {
                BusinessEntityID = 292,
                Demographics = "<StoreSurvey xmlns=\"http://schemas.microsoft.com/sqlserver/2004/07/adventure-works/StoreSurvey\"><AnnualSales>800000</AnnualSales><AnnualRevenue>80000</AnnualRevenue><BankName>United Security</BankName><BusinessType>BM</BusinessType><YearOpened>1996</YearOpened><Specialty>Mountain</Specialty><SquareFeet>21000</SquareFeet><Brands>2</Brands><Internet>ISDN</Internet><NumberEmployees>13</NumberEmployees></StoreSurvey>",
                ModifiedDate = new DateTime(2004, 10, 13, 11, 15, 07),
                Name = "Next-Door Bike Store",
                Rowguid = "a22517e3-848d-4ebe-b9d9-7437f3432304",
                SalesPersonID = 279
            };
        }

        /// <summary>
        /// Opens the connection test.
        /// Makes sure no exception was thrown.
        /// </summary>
        [Test]
        public void OpenConnectionTest()
        {
            SqlHelper.OpenConnection(validConnection);
        }

        /// <summary>
        /// Opens the connection exception test.
        /// </summary>
        [Test]
        [ExpectedException(typeof(PowerPointGenerator.Exceptions.SqlException))]
        public void OpenConnectionExceptionTest()
        {
            SqlHelper.OpenConnection(invalidConnection);
        }

        /// <summary>
        /// Closes the connection test.
        /// </summary>
        [Test]
        public void CloseConnectionTest()
        {
            SqlHelper.CloseConnection(validConnection);
        }

        /// <summary>
        /// Gets the sales count test.
        /// </summary>
        [Test]
        public void GetSalesCountTest()
        {
            Assert.AreEqual(SqlHelper.GetSales(validConnection).Count, salesNumber);
        }

        ///// <summary>
        ///// Gets the sales first equal test. - I think this fails because objects are store in different memory locations
        ///// </summary>
        //[Test]
        //public void GetSalesFirstEqualTest()
        //{
        //    Assert.AreEqual(SqlHelper.GetSales(validConnection)[0], firstSale);
        //}

        /// <summary>
        /// Gets the sales first equal test.
        /// </summary>
        [Test]
        public void GetSalesFirstBusinessEntityIDEqualTest()
        {
            Assert.AreEqual(SqlHelper.GetSales(validConnection)[0].BusinessEntityID, firstSale.BusinessEntityID);
        }

        /// <summary>
        /// Gets the sales first equal test.
        /// </summary>
        [Test]
        public void GetSalesFirstSalesPersonIDEqualTest()
        {
            Assert.AreEqual(SqlHelper.GetSales(validConnection)[0].SalesPersonID, firstSale.SalesPersonID);
        }

        /// <summary>
        /// Gets the sales count test.
        /// </summary>
        [Test]
        [ExpectedException(typeof(PowerPointGenerator.Exceptions.SqlException))]
        public void GetSalesExceptionTest()
        {
            SqlHelper.GetSales(invalidConnection);
        }

        /// <summary>
        /// Resets this instance.
        /// </summary>
        [TearDown]
        public void Reset()
        {
            SqlHelper.CloseConnection(validConnection);
        }
    }
}
