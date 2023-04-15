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
        /// Handles the Running of a <see cref="GenericReport"/> Synchronously
        /// </summary>
        /// <param name="reportToRun">The <see cref="GenericReport"/> to run</param>
        /// <returns>The resulting <see cref="ReportResultSet"/></returns>
        /// <exception cref="ArgumentException">Throws an <see cref="ArgumentException"/> if the <see cref="GenericReport"/> is Invalid</exception>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Connection is not established / setup.</exception>
        public static ReportResultSet RunSingleReport(GenericReport reportToRun)
        {
            var type = reportToRun is SqlReport ? 
                ReportType.SqlReport : 
                ReportType.StoredProcedure;

            var commandText = string.Empty;
            var commandType = type == ReportType.SqlReport ? 
                CommandType.Text :
                CommandType.StoredProcedure;

            if (type is ReportType.SqlReport)
            {
                var report = reportToRun as SqlReport;
                if (!report.IsValid())
                {
                    throw new ArgumentException(SqlRepository._invalidReportMessage);
                }

                commandText = report.Query;
            }
            else
            {
                var report = reportToRun as StoredProcedureReport;
                if (!report.IsValid())
                {
                    throw new ArgumentException(SqlRepository._invalidReportMessage);
                }

                commandText = report.ProcedureName;
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
                command.CommandText = commandText;
                command.CommandType = commandType;
                command.AddParamsIfRequired(reportToRun);

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
                    resultSet.ReportUsed = reportToRun;

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
        /// Handles the Running of a <see cref="GenericReport"/> Asynchronously
        /// </summary>
        /// <param name="reportToRun">The <see cref="GenericReport"/> to run</param>
        /// <returns>The resulting <see cref="ReportResultSet"/></returns>
        /// <exception cref="ArgumentException">Throws an <see cref="ArgumentException"/> if the <see cref="GenericReport"/> is Invalid</exception>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Connection is not established / setup.</exception>
        public static async Task<ReportResultSet> RunSingleReportAsync(GenericReport reportToRun)
        {
            var type = reportToRun is SqlReport ?
                ReportType.SqlReport :
                ReportType.StoredProcedure;

            var commandText = string.Empty;
            var commandType = type == ReportType.SqlReport ?
                CommandType.Text :
                CommandType.StoredProcedure;

            if (type is ReportType.SqlReport)
            {
                var report = reportToRun as SqlReport;
                if (!report.IsValid())
                {
                    throw new ArgumentException(SqlRepository._invalidReportMessage);
                }

                commandText = report.Query;
            }
            else
            {
                var report = reportToRun as StoredProcedureReport;
                if (!report.IsValid())
                {
                    throw new ArgumentException(SqlRepository._invalidReportMessage);
                }

                commandText = report.ProcedureName;
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
                command.CommandText = commandText;
                command.CommandType = commandType;
                command.AddParamsIfRequired(reportToRun);

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
                    resultSet.ReportUsed = reportToRun;

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
        /// Creates a new Instance of the <see cref="Stopwatch"/> class and Starts it
        /// </summary>
        private static void StartStopwatch()
        {
            SqlRepository._sw = new Stopwatch();
            SqlRepository._sw.Start();
        }
    }
}