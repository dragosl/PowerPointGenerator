using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointGenerator.Model
{
    /// <summary>
    /// Customer model class - holds data of a customer.
    /// </summary>
    public class Customer
    {
        //ID name address money

        /// <summary>
        /// Gets or sets the ID of this customer.
        /// </summary>
        /// <value>
        /// The ID.
        /// </value>
        public int ID { get; set; }

        /// <summary>
        /// Gets or sets this customer's name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the address of this customer.
        /// </summary>
        /// <value>
        /// The address.
        /// </value>
        public string Address { get; set; }

        /// <summary>
        /// Gets or sets the money of this customer.
        /// </summary>
        /// <value>
        /// The money.
        /// </value>
        public int Money { get; set; }
    }
}
