using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using PowerPointGenerator.Model;

namespace PowerPointGenerator.Helpers
{
    /// <summary>
    /// Provides database functionality.
    /// </summary>
    public static class SqlHelper
    {
        /// <summary>
        /// Open Database Connection
        /// </summary>
        /// <param name="connection">The connection.</param>
        /// <exception cref="Exceptions.SqlException"></exception>
        public static void OpenConnection(SqlConnection connection)
        {
            try
            {
                connection.Open();
            }
            catch (Exception ex)
            {
                throw new Exceptions.SqlException(ex);
            }
        }

        /// <summary>
        /// Close Database connection
        /// </summary>
        /// <param name="connection">The connection.</param>
        /// <exception cref="Exceptions.SqlException"></exception>
        public static void CloseConnection(SqlConnection connection)
        {
            try
            {
                connection.Close();
            }
            catch (Exception ex)
            {
                throw new Exceptions.SqlException(ex);
            }
        }

        /// <summary>
        /// Gets all the sales from the database.
        /// </summary>
        /// <param name="connection">The database connection.</param>
        /// <returns>The obtained sales.</returns>
        public static List<Sale> GetSales(SqlConnection connection)
        {
            string sql;
            try
            {
                OpenConnection(connection);
                List<Sale> sales = new List<Sale>();
                sql = "SELECT * FROM Sales.Store;";

                SqlCommand command = new SqlCommand(sql, connection);
                SqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {
                    Sale sale = new Sale
                    {
                        BusinessEntityID = (int)dr["BusinessEntityID"],
                        Name = dr["Name"].ToString(),
                        SalesPersonID = (int)dr["SalesPersonID"],
                        Demographics = dr["Demographics"].ToString(),
                        Rowguid = dr["Rowguid"].ToString(),
                        ModifiedDate = (DateTime)dr["ModifiedDate"]
                    };

                    sales.Add(sale);
                }

                dr.Close();
                return sales;
            }
            catch (Exception ex)
            {
                throw new Exceptions.SqlException(ex);
            }
            finally
            {
                CloseConnection(connection);
            }
        }
    }
}
