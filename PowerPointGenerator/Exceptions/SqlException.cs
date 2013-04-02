using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointGenerator.Exceptions
{
    /// <summary>
    /// Exceptions defined for sql
    /// </summary>
    public class SqlException : Exception
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlException"/> class.
        /// </summary>
        /// <param name="ex">exception occurred</param>
        public SqlException(Exception ex)
            : base("An exception has occurred. Please check the inner exception for details.", ex)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlException"/> class.
        /// </summary>
        /// <param name="message">exception message</param>
        /// <param name="ex">exception occurred</param>
        public SqlException(string message, Exception ex)
            : base(message, ex)
        {
        }

        #endregion Constructors
    }
}
