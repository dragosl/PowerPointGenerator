using System;
using System.IO;

namespace PowerPointGenerator.Helpers
{
    public static class StreamHelper
    {
        public static Stream GenerateRandomStream()
        {
            byte[] array = new byte[899];
            // Use Random class and NextBytes method.
            // ... Display the bytes with following method.
            Random random = new Random();
            random.NextBytes(array);
            var stream = new MemoryStream(array);
            stream.Position = 0;
            return stream;
        }
    }
}
