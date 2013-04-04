using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using PowerPointGenerator.Helpers;

namespace UnitTests.PowerPointGeneratorTests
{
    [TestFixture]
    public class StreamHelperTest
    {
        [Test]
        public void StreamGenerateReturnNotNullTest()
        {
            Assert.IsTrue(StreamHelper.GenerateRandomStream() != null);
        }
    }
}
