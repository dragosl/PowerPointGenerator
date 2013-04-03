using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointGenerator.Model
{
    /// <summary>
    /// Sale model class - holds data of a sale
    /// </summary>
    public class Sale
    {
        /// <summary>
        /// Gets or sets the ID of this sale.
        /// </summary>
        /// <value>
        /// The ID.
        /// </value>
        public int BusinessEntityID { get; set; }

        /// <summary>
        /// Gets or sets the number of items of the sale.
        /// </summary>
        /// <value>
        /// The pieces.
        /// </value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the body of this sale.
        /// </summary>
        /// <value>
        /// The body.
        /// </value>
        public int SalesPersonID { get; set; }

        /// <summary>
        /// Gets or sets the price of the sale.
        /// </summary>
        /// <value>
        /// The price.
        /// </value>
        public string Demographics { get; set; }        

        /// <summary>
        /// Gets or sets the order to which this sale belongs.
        /// </summary>
        /// <value>
        /// The order.
        /// </value>
        public string Rowguid { get; set; }

        /// <summary>
        /// Gets or sets the date when this sale was updated.
        /// </summary>
        /// <value>
        /// The updated.
        /// </value>
        public DateTime ModifiedDate { get; set; }

        public override string ToString()
        {
            return this.BusinessEntityID + " " + this.Name + " " + this.SalesPersonID
                + " " + this.Demographics + " " + this.Rowguid + " " + this.ModifiedDate;
        }
    }
}
