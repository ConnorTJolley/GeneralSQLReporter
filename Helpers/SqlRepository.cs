namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.Common;
    using System.Data.SqlClient;
    using System.Threading.Tasks;
    using GeneralSQLReporter.Models;

    public static class SqlRepository
    {
        /// <summary>
        /// Private instance of the <see cref="SqlConnection"/> that the <see cref="SqlRepository"/> will use.
        /// </summary>
        private static readonly SqlConnection _conn;

        /// <summary>
        /// Initializes the Static instance of the <see cref="SqlRepository"/>
        /// </summary>
        static SqlRepository()
        {
            SqlRepository._conn = new SqlConnection("");
        }

        /// <summary>
        /// Gets the current <see cref="SqlConnection"/> that is setup for the <see cref="SqlRepository"/> to use.
        /// </summary>
        /// <returns>The <see cref="SqlConnection"/> setup.</returns>
        public static SqlConnection GetConnection() => SqlRepository._conn;

        /// <summary>
        /// Synchronously sets up the <see cref="SqlConnection"/> for the <see cref="SqlRepository"/> to use.
        /// 
        /// This must be done before attempting to set the Query or Running any reports
        /// </summary>
        /// <param name="connectionString">The Connection String for the SQL Server</param>
        /// <returns>
        /// <see cref="true"/> if successfully setup and connected.
        /// <see cref="false"/> if not
        /// </returns>
        public static bool SetupSqlConnection(string connectionString)
        {
            SqlRepository._conn.ConnectionString = connectionString;
            return SqlRepository._conn.CheckConnection();
        }

        /// <summary>
        /// Asynchronously sets up the <see cref="SqlConnection"/> for the <see cref="SqlRepository"/> to use.
        /// 
        /// This must be done before attempting to set the Query or Running any reports
        /// </summary>
        /// <param name="connectionString">The Connection String for the SQL Server</param>
        /// <returns>
        /// <see cref="true"/> if successfully setup and connected.
        /// <see cref="false"/> if not
        /// </returns>
        public static async Task<bool> SetupSqlConnectionAsync(string connectionString)
        {
            SqlRepository._conn.ConnectionString = connectionString;
            return await SqlRepository._conn.CheckConnectionAsync();
        }

        public static ReportResultSet RunSingleReport(SqlReport report)
        {
            if (!report.IsValid())
            {
                throw new ArgumentException("Report as Invalid, please check all required values are set " +
                    "and that SmtpEmailSender is configured if required.");
            }

            if (!SqlRepository._conn.CheckConnection())
            {
                throw new Exception("SqlRepository Connection was not able to Connect, " +
                    "please check Connection has been setup.");
            }

            try
            {
                var resultSet = new ReportResultSet();
                var command = SqlRepository._conn.CreateCommand();
                command.CommandText = report.Query;

                using (var schemaOnlyReader = command.ExecuteReader(CommandBehavior.SchemaOnly))
                {
                    var schema = schemaOnlyReader.GetSchemaTable();
                    var rows = schema.Rows;

                    var colIndex = 0;
                    foreach (DataRow col in rows)
                    {
                        var name = col.Field<string>("ColumnName");

                        resultSet.Columns.Add(new SqlColumn
                        {
                            Index = colIndex,
                            Name = name
                        });

                        colIndex++;
                    }
                }

                using (var reader = command.ExecuteReader())
                {
                    resultSet.ReportUsed = report;

                    var rowIndex = 0;

                    while (reader.Read())
                    {
                        var sqlRow = new SqlRow(rowIndex);

                        for (var i = 0; i < reader.FieldCount; i++)
                        {
                            sqlRow.Values.Add(new SqlRowValue
                            {
                                ColumnIndex = i,
                                Value = reader.GetValue(i),
                                DataType = reader.GetFieldType(i),
                                RowNumber = rowIndex
                            });
                        }

                        resultSet.Rows.Add(sqlRow);
                        rowIndex++;
                    }
                }

                command.Dispose();

                return resultSet;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error Running Report. Ex: {ex.Message}");
                throw ex;
            }
        }

        public static async Task<ReportResultSet> RunSingleReportAsync(SqlReport report)
        {
            if (!report.IsValid())
            {
                throw new ArgumentException("Report as Invalid, please check all required values are set " +
                    "and that SmtpEmailSender is configured if required.");
            }

            if (!await SqlRepository._conn.CheckConnectionAsync())
            {
                throw new Exception("SqlRepository Connection was not able to Connect, " +
                    "please check Connection has been setup.");
            }

            try
            {
                var resultSet = new ReportResultSet();
                var command = SqlRepository._conn.CreateCommand();
                command.CommandText = report.Query;

                using (var schemaOnlyReader = command.ExecuteReader(CommandBehavior.SchemaOnly))
                {
                    var schema = schemaOnlyReader.GetSchemaTable();
                    var rows = schema.Rows;

                    var colIndex = 0;
                    foreach (DataRow col in rows)
                    {
                        var name = col.Field<string>("ColumnName");

                        resultSet.Columns.Add(new SqlColumn
                        {
                            Index = colIndex,
                            Name = name
                        });

                        colIndex++;
                    }
                }

                using (var reader = await command.ExecuteReaderAsync())
                {
                    resultSet.ReportUsed = report;

                    var rowIndex = 0;

                    while (await reader.ReadAsync())
                    {
                        var sqlRow = new SqlRow(rowIndex);

                        for (var i = 0; i < reader.FieldCount; i++)
                        {
                            sqlRow.Values.Add(new SqlRowValue
                            {
                                ColumnIndex = i,
                                Value = reader.GetValue(i),
                                DataType = reader.GetFieldType(i),
                                RowNumber = rowIndex
                            });
                        }

                        resultSet.Rows.Add(sqlRow);
                        rowIndex++;
                    }
                }

                command.Dispose();

                return resultSet;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error Running Report. Ex: {ex.Message}");
                throw ex;
            }
        }
    }
}