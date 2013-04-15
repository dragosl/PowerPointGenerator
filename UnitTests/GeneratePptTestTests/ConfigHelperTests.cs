using GeneratePptTest.Helpers;
using NUnit.Framework;

namespace UnitTests.GeneratePptTestTests
{
    [TestFixture]
    public class ConfigHelperTests
    {
        string postgreConnectionString;

        string microsoftConnectionString;

        const string FilePath = "@Settings.ini";

        /// <summary>
        /// Inits this instance.
        /// </summary>
        [SetUp]
        public void Init()
        {
            postgreConnectionString = "Server=127.0.0.1;Port=5432;User Id=postgres;Database=sales;Password=postgres;";

            microsoftConnectionString = "server=DRAGOSL-PC\\SQLEXPRESS;database=AdventureWorks2008;Trusted_Connection=True;";
        }

        /// <summary>
        /// Generate postgre connection string equal test.
        /// </summary>
        [Test]
        public void GeneratePostgreConnectionStringEqualTest()
        {
            Assert.AreEqual(ConfigHelper.GenerateConnectionString(), postgreConnectionString);
        }

        /// <summary>
        /// Generate micrososft connection string equal test.
        /// </summary>
        [Test]
        public void GenerateMicrososftConnectionStringEqualTest()
        {
            Assert.AreEqual(ConfigHelper.GenerateConnectionStringMssql(), microsoftConnectionString);
        }

        /// <summary>
        /// Generate postgre connection string from file equal test.
        /// </summary>
        [Test]
        public void GeneratePostgreConnectionStringFromFileEqualTest()
        {
            Assert.AreEqual(ConfigHelper.GenerateConnectionStringFromFile(FilePath), postgreConnectionString);
        }

        /// <summary>
        /// Generates the micrososft connection string from file equal test.
        /// </summary>
        [Test]
        public void GenerateMicrososftConnectionStringFromFileEqualTest()
        {
            Assert.AreEqual(ConfigHelper.GenerateConnectionStringMssqlFromFile(FilePath), microsoftConnectionString);
        }
    }
}
