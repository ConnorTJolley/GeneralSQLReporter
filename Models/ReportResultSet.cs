namespace GeneralSQLReporter.Models
{
    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using GeneralSQLReporter.Helpers;
    using Newtonsoft.Json;

    /// <summary>
    /// Interaction Logic for the <see cref="ReportResultSet"/> Class.
    /// </summary>
    /// <remarks>
    /// This object is returned by any of the "RunSingleReport" / "RunSingleReportAsync" methods on the <see cref="SqlRepository"/> Helper class.
    /// 
    /// This object is also used by the <see cref="SqlReportExporter"/> Helper class that processes and generates the output.
    /// </remarks>
    public class ReportResultSet
    {
        /// <summary>
        /// Gets the <see cref="GenericReport"/> that was Used in the execution of the Report.
        /// </summary>
        /// <remarks>
        /// Can be of type <see cref="SqlReport"/> or <see cref="StoredProcedureReport"/>
        /// </remarks>
        public GenericReport ReportUsed { get; internal set; }

        /// <summary>
        /// Gets the <see cref="List{SqlColumn}"/> that the Query Returned.
        /// </summary>
        /// <remarks>
        /// This is auto-populated once Running the Report via the Schema gathered by the <see cref="SqlDataReader"/>
        /// </remarks>
        public List<SqlColumn> Columns { get; internal set; } = new List<SqlColumn>();

        /// <summary>
        /// Gets the <see cref="List{SqlRow}"/> that the Query Returned
        /// </summary>
        public List<SqlRow> Rows { get; set; } = new List<SqlRow>();

        /// <summary>
        /// Converts the <see cref="ReportResultSet"/> Object into its JSON representation
        /// </summary>
        /// <returns>The JSON representation of the <see cref="ReportResultSet"/> Object</returns>
        public string ToJson() => JsonConvert.SerializeObject(this);

        /// <summary>
        /// Gets the Elapsed Time in Ms that the <see cref="ReportResultSet.ReportUsed"/> took to Run
        /// </summary>
        public long ElapsedTimeInMs { get; internal set; }

        /// <summary>
        /// Handles Converting the <see cref="ReportResultSet"/> into a SQL Query Representation.
        /// </summary>
        /// <returns>The SQL Query that was generated and used to run the Report.</returns>
        public string ToSqlQuery()
        {
            if (this.ReportUsed == null)
            {
                return "-- Unable to determine Report Type.";
            }

            var fullQueryString = string.Empty;

            switch (this.GetReportType())
            {
                case ReportType.SqlReport:
                    fullQueryString += this.ReportUsed.Parameters.ToSqlParams();
                    fullQueryString += (this.ReportUsed as SqlReport).Query;
                    break;

                case ReportType.StoredProcedure:
                    var procedure = this.ReportUsed as StoredProcedureReport;
                    var tab = "\t";

                    var baseQuery = $"EXEC {tab} {procedure.ProcedureName}{Environment.NewLine}";
                    procedure.Parameters.ForEach(p =>
                    {
                        baseQuery += $"{tab}{p.ParameterName} = {p.ToQueryString()},{Environment.NewLine}";
                    });

                    var trimmed = baseQuery.Trim();
                    trimmed = trimmed.Remove(trimmed.Length - 1);
                    fullQueryString += trimmed;
                    break;
            }

            return fullQueryString;
        }

        /// <summary>
        /// Handles Determining the <see cref="ReportType"/> that was used via <see cref="ReportResultSet.ReportUsed"/>
        /// </summary>
        /// <returns>The determined <see cref="ReportType"/> if it was found</returns>
        public ReportType GetReportType()
        {
            if (this.ReportUsed == null)
            {
                return ReportType.Unknown;
            }

            return this.ReportUsed is SqlReport ? 
                ReportType.SqlReport : 
                ReportType.StoredProcedure;
        }

        /// <summary>
        /// Gets a more structured <see cref="string"/> Representation of the <see cref="ReportResultSet"/> object.
        /// </summary>
        /// <returns>The formatted string containing: 
        /// - The Type of Report
        /// - The Columns and Indexes
        /// - The Record Count
        /// - The Query / Stored Procedure Used 
        /// - The Parameters if any
        /// </returns>
        /// <example>
        /// 
        /// StoredProcedure Results:
        /// Columns and Index: 
        /// [Id (0) | F_Name (1) | L_Name (2)]
        /// Record Count: 482,000
        /// Stored Procedure Used: sp_GetAllNames
        /// Elapsed Time: 450ms
        /// </example>
        public override string ToString()
        {
            var repUsed = this.GetReportType();

            var newLine = Environment.NewLine;
            var baseMsg = $"{repUsed} Results:{newLine}";

            var colString = string.Empty;
            this.Columns.ForEach(col =>
            {
                colString += $"{col.Name} ({col.Index}) | ";
            });

            //// Remove trailing pipe
            var trimmedColString = colString.Trim();
            var lastIndex = trimmedColString.LastIndexOf('|');
            colString = trimmedColString.Remove(lastIndex, 1).Trim();

            baseMsg += $"Columns & Index: {newLine}[{colString}]{newLine}";
            baseMsg += $"Record Count: {this.Rows.Count:N0}{newLine}";

            var typeMsg = repUsed == ReportType.SqlReport ? "Query" : "Stored Procedure";
            baseMsg += $"{typeMsg} Used: {newLine}";

            if (repUsed == ReportType.SqlReport)
            {
                var _ = this.ReportUsed as SqlReport;
                baseMsg += $"{_.Query}{newLine}";
                baseMsg += _.Parameters.ToFormattedString();
            }
            else
            {
                var _ = this.ReportUsed as StoredProcedureReport;
                baseMsg += $"{_.ProcedureName}{newLine}";
                baseMsg += _.Parameters.ToFormattedString();
            }

            baseMsg += $"{newLine}Elapsed Time: {this.ElapsedTimeInMs}ms";

            return baseMsg;
        }
    }
}