using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GeneratePptTest.Business;
using NUnit.Framework;

namespace UnitTests.GeneratePptTestTests
{
    [TestFixture]
    public class Form1ManagerTests
    {
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
            templatePath = @"Templates\template.ppt";
            savePptFilePath = @"demoppt.ppt";

            invalidTemplatePath = @"Templates\template.pptx";
        }

        /// <summary>
        /// Test which verifies if the ppt was generated with success.
        /// </summary>
        [Test]
        public void GeneratePptTest()
        {
            Assert.IsTrue(Form1Manager.GeneratePpt(templatePath, savePptFilePath));
        }

        /// <summary>
        /// Test which verifies if the ppt failed to be generated, because of the template inconsistency.
        /// </summary>
        [Test]
        public void GeneratePptInvalidTemplateTest()
        {
            Assert.IsFalse(Form1Manager.GeneratePpt(invalidTemplatePath, savePptFilePath));
        }
    }
}
