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
