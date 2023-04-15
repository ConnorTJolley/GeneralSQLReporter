namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.Data;
    using System.Data.SqlClient;
    using System.Diagnostics;
    using System.Linq;
    using System.Threading.Tasks;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Interaction Logic for the Helper Class of <see cref="SqlRepository"/>
    /// </summary>
    /// <remarks>
    /// This Helper class handles the Connection, Querying and Creation of the <see cref="ReportResultSet"/> via either a 
    /// <see cref="SqlReport"/> or <see cref="StoredProcedureReport"/>.
    /// </remarks>
    public static class SqlRepository
    {
        /// <summary>
        /// Instance of the Invalid Report Error Message
        /// </summary>
        private static readonly string _invalidReportMessage = "Report is Invalid, please chell all required values are set " +
                    "and that SmtpEmailSender is configured if required.";

        /// <summary>
        /// Instance of the Connection Issue Error Message
        /// </summary>
        private static readonly string _connectionIssueMessage = "SqlRepository Connection was not able to Connect, " +
                    "please check Connection has been setup.";

        /// <summary>
        /// Instance of the <see cref="Stopwatch"/> for Tracking Report Runtime
        /// </summary>
        private static Stopwatch _sw;

        /// <summary>
        /// Private instance of the <see cref="SqlConnection"/> that the <see cref="SqlRepository"/> will use.
        /// </summary>
        private static readonly SqlConnection _conn = new SqlConnection();

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
        /// True if successfully setup and connected.
        /// False if not
        /// </returns>
        public static bool SetupSqlConnection(string connectionString)
        {
            if (SqlRepository._conn.State == ConnectionState.Open)
            {
                //// Open connection, close before changing ConnectionString
                SqlRepository._conn.Close();
            }

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
        /// True if successfully setup and connected.
        /// False if not
        /// </returns>
        public static async Task<bool> SetupSqlConnectionAsync(string connectionString)
        {
            if (SqlRepository._conn.State == ConnectionState.Open)
            {
                //// Open connection, close before changing ConnectionString
                SqlRepository._conn.Close();
            }

            SqlRepository._conn.ConnectionString = connectionString;
            return await SqlRepository._conn.CheckConnectionAsync();
        }

        /// <summary>
        /// Runs a Single Report of Type <see cref="StoredProcedureReport"/> For Stored Procedures Sychonrously
        /// </summary>
        /// <param name="report">The <see cref="StoredProcedureReport"/> to run</param>
        /// <returns>The <see cref="ReportResultSet"/></returns>
        /// <exception cref="ArgumentException">Throws an <see cref="ArgumentException"/> if the Report is not valid.</exception>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Connection is not established / setup.</exception>
        public static ReportResultSet RunSingleReport(StoredProcedureReport report)
        {
            if (!report.IsValid())
            {
                throw new ArgumentException(SqlRepository._invalidReportMessage);
            }

            if (!SqlRepository._conn.CheckConnection())
            {
                throw new Exception(SqlRepository._connectionIssueMessage);
            }

            try
            {
                SqlRepository.StartStopwatch();

                var resultSet = new ReportResultSet();
                var query = report.ProcedureName;
                var command = SqlRepository._conn.CreateCommand();
                command.CommandText = query;
                command.CommandType = CommandType.StoredProcedure;
                command.AddParamsIfRequired(report);

                SqlRepository._conn.CheckConnection();

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

                SqlRepository._sw.Stop();
                resultSet.ElapsedTimeInMs = SqlRepository._sw.ElapsedMilliseconds;

                return resultSet;
            }
            catch (Exception ex)
            {
                SqlRepository._sw.Stop();
                Console.WriteLine($"Error Running Report. Ex: {ex.Message}");
                throw ex;
            }
        }

        /// <summary>
        /// Runs a Single Report of Type <see cref="StoredProcedureReport"/> For Stored Procedures Asynchronously
        /// </summary>
        /// <param name="report">The <see cref="StoredProcedureReport"/> to run</param>
        /// <returns>The <see cref="ReportResultSet"/></returns>
        /// <exception cref="ArgumentException">Throws an <see cref="ArgumentException"/> if the Report is not valid.</exception>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Connection is not established / setup.</exception>
        public static async Task<ReportResultSet> RunSingleReportAsync(StoredProcedureReport report)
        {
            if (!report.IsValid())
            {
                throw new ArgumentException(SqlRepository._invalidReportMessage);
            }

            if (!SqlRepository._conn.CheckConnection())
            {
                throw new Exception(SqlRepository._connectionIssueMessage);
            }

            try
            {
                SqlRepository.StartStopwatch();

                var resultSet = new ReportResultSet();
                var query = report.ProcedureName;
                var command = SqlRepository._conn.CreateCommand();
                command.CommandText = query;
                command.CommandType = CommandType.StoredProcedure;
                command.AddParamsIfRequired(report);

                if (report.Parameters.Any())
                {
                    report.Parameters.ForEach(param =>
                    {
                        command.Parameters.Add(param);
                    });
                }

                await SqlRepository._conn.CheckConnectionAsync();

                using (var schemaOnlyReader = await command.ExecuteReaderAsync(CommandBehavior.SchemaOnly))
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

                SqlRepository._sw.Stop();
                resultSet.ElapsedTimeInMs = SqlRepository._sw.ElapsedMilliseconds;

                return resultSet;
            }
            catch (Exception ex)
            {
                SqlRepository._sw.Stop();
                Console.WriteLine($"Error Running Report. Ex: {ex.Message}");
                throw ex;
            }
        }

        /// <summary>
        /// Runs a Single Report of Type <see cref="SqlRepository"/> For Query Reports
        /// </summary>
        /// <param name="report">The <see cref="SqlRepository"/> to run</param>
        /// <returns>The <see cref="ReportResultSet"/></returns>
        /// <exception cref="ArgumentException">Throws an <see cref="ArgumentException"/> if the Report is not valid.</exception>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Connection is not established / setup.</exception>
        public static ReportResultSet RunSingleReport(SqlReport report)
        {
            if (!report.IsValid())
            {
                throw new ArgumentException(SqlRepository._invalidReportMessage);
            }

            if (!SqlRepository._conn.CheckConnection())
            {
                throw new Exception(SqlRepository._connectionIssueMessage);
            }

            try
            {
                SqlRepository.StartStopwatch();

                var resultSet = new ReportResultSet();
                var command = SqlRepository._conn.CreateCommand();
                command.CommandText = report.Query;
                command.AddParamsIfRequired(report);

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

                SqlRepository._sw.Stop();
                resultSet.ElapsedTimeInMs = SqlRepository._sw.ElapsedMilliseconds;

                return resultSet;
            }
            catch (Exception ex)
            {
                SqlRepository._sw.Stop();
                Console.WriteLine($"Error Running Report. Ex: {ex.Message}");
                throw ex;
            }
        }

        /// <summary>
        /// Runs a Single Report of Type <see cref="SqlRepository"/> For Query Reports Asynchronously
        /// </summary>
        /// <param name="report">The <see cref="SqlRepository"/> to run</param>
        /// <returns>The <see cref="ReportResultSet"/></returns>
        /// <exception cref="ArgumentException">Throws an <see cref="ArgumentException"/> if the Report is not valid.</exception>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Connection is not established / setup.</exception>
        public static async Task<ReportResultSet> RunSingleReportAsync(SqlReport report)
        {
            if (!report.IsValid())
            {
                throw new ArgumentException(SqlRepository._invalidReportMessage);
            }

            if (!await SqlRepository._conn.CheckConnectionAsync())
            {
                throw new Exception(SqlRepository._connectionIssueMessage);
            }

            try
            {
                SqlRepository.StartStopwatch();

                var resultSet = new ReportResultSet();
                var command = SqlRepository._conn.CreateCommand();
                command.CommandText = report.Query;
                command.AddParamsIfRequired(report);

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

                SqlRepository._sw.Stop();
                resultSet.ElapsedTimeInMs = SqlRepository._sw.ElapsedMilliseconds;

                return resultSet;
            }
            catch (Exception ex)
            {
                SqlRepository._sw.Stop();
                Console.WriteLine($"Error Running Report. Ex: {ex.Message}");
                throw ex;
            }
        }

        /// <summary>
        /// Creates a new Instance of the <see cref="Stopwatch"/> class and Starts it
        /// </summary>
        private static void StartStopwatch()
        {
            SqlRepository._sw = new Stopwatch();
            SqlRepository._sw.Start();
        }
    }
}