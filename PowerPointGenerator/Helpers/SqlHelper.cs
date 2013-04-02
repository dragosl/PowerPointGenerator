using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;
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
        public static void OpenConnection(NpgsqlConnection connection)
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
        public static void CloseConnection(NpgsqlConnection connection)
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
        public static List<Sale> GetSales(NpgsqlConnection connection)
        {
            string sql = string.Empty;
            try
            {
                OpenConnection(connection);
                List<Sale> sales = new List<Sale>();
                sql = "SELECT * FROM sale;";

                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {
                    Sale sale = new Sale()
                    {
                        ID = (int)dr["id"],
                        Body=dr["body"].ToString(),
                        Order=(int)dr["order"],
                        Pieces=(int)dr["pieces"],
                        Price=(int)dr["price"],
                        Product=(int)dr["product"],
                        Updated=(DateTime)dr["updated"]
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
