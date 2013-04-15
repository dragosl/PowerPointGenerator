using System.Collections.Generic;
using System.Data.SqlClient;
using PowerPointGenerator.Helpers;
using PowerPointGenerator.Model;

namespace PowerPointGenerator.Managers
{
    /// <summary>
    /// Provides business for the store.
    /// </summary>
    public class StoreManager
    {
        #region Properties

        /// <summary>
        /// MSSQL Connection
        /// </summary>
        private SqlConnection Connection { get; set; }

        /// <summary>
        /// Gets or sets the categories of the store.
        /// </summary>
        /// <value>
        /// The categories.
        /// </value>
        public List<Sale> Sales { get; set; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="StoreManager" /> class.
        /// </summary>
        public StoreManager(string connectionString)
        {
            this.Connection = new SqlConnection(connectionString);
            this.Sales = SqlHelper.GetSales(this.Connection);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="StoreManager" /> class.
        /// </summary>
        /// <param name="sales">The sales.</param>
        public StoreManager(List<Sale> sales)
        {
            this.Sales = sales;
        }

        #endregion Constructors

        #region Public methods

        /// <summary>
        /// Generates the Ppt file based on a given template.
        /// </summary>
        /// <param name="templatePath">The pptx template path.</param>
        /// <param name="exportPptFilePath">The export PPT file path.</param>
        /// <returns></returns>
        public bool GeneratePpt(string templatePath, string exportPptFilePath)
        {
            // use a ppt helper to return a ppt object with sales(or just dialog to save it)                
            bool x = System.IO.File.Exists(templatePath);
            System.Diagnostics.Debug.Assert(x);

            return PptHelper.InsertSalesInTemplate(this.Sales, templatePath, exportPptFilePath);
        }

        #endregion Public methods
    }
}
