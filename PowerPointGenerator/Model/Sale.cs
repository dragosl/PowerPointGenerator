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
        public int ID { get; set; }

        /// <summary>
        /// Gets or sets the number of items of the sale.
        /// </summary>
        /// <value>
        /// The pieces.
        /// </value>
        public int Pieces { get; set; }

        /// <summary>
        /// Gets or sets the body of this sale.
        /// </summary>
        /// <value>
        /// The body.
        /// </value>
        public string Body { get; set; }

        /// <summary>
        /// Gets or sets the price of the sale.
        /// </summary>
        /// <value>
        /// The price.
        /// </value>
        public int Price { get; set; }

        /// <summary>
        /// Gets or sets the date when this sale was updated.
        /// </summary>
        /// <value>
        /// The updated.
        /// </value>
        public DateTime Updated { get; set; }

        /// <summary>
        /// Gets or sets the order to which this sale belongs.
        /// </summary>
        /// <value>
        /// The order.
        /// </value>
        public int Order { get; set; }

        /// <summary>
        /// Gets or sets the product which was sold.
        /// </summary>
        /// <value>
        /// The product.
        /// </value>
        public int Product { get; set; }

        public override string ToString()
        {
            return this.Body + " " + this.ID + " " + this.Order
                + " " + this.Pieces + " " + this.Price + " " + this.Product + " " + this.Updated;
        }
    }
}
