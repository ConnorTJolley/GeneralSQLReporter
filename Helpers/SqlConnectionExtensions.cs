namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.Data.SqlClient;
    using System.Data;
    using System.Threading.Tasks;

    public static class SqlConnectionExtensions
    {
        /// <summary>
        /// Handles the Checking of a Connection Asynchronously.
        /// </summary>
        /// <param name="conn">The <see cref="SqlConnection"/> to check</param>
        /// <returns>True if Connected, False if an <see cref="Exception"/> was encountered</returns>
        public static async Task<bool> CheckConnectionAsync(this SqlConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                {
                    return true;
                }

                await conn.OpenAsync();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Handles the Checking of a Connection Synchronously.
        /// </summary>
        /// <param name="conn">The <see cref="SqlConnection"/> to check</param>
        /// <returns>True if Connected, False if an <see cref="Exception"/> was encountered</returns>
        public static bool CheckConnection(this SqlConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                {
                    return true;
                }

                conn.Open();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
    }
}