namespace GeneralSQLReporter.Models
{
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using GeneralSQLReporter.Enums;
    using GeneralSQLReporter.Helpers;

    /// <summary>
    /// Interaction Logic for the <see cref="StoredProcedureReport"/> Class
    /// </summary>
    /// <remarks>
    /// Provides important properties for being able to create a "generic" Stored Procedure based Report with or without 
    /// using <see cref="SqlParameter"/>'s against whatever table is defined in the <see cref="SqlRepository"/>.
    /// </remarks>
    public class StoredProcedureReport : GenericReport
    {
        /// <summary>
        /// Gets the Name of the Stored Procedure
        /// </summary>
        public string ProcedureName { get; internal set; }

        /// <summary>
        /// Creates a new Instance of the <see cref="StoredProcedureReport"/> Class
        /// </summary>
        /// <param name="procedureName">The Name of the Stored Procedure to run</param>
        /// <param name="parameters">The <see cref="List{SqlParameter}"/> to add to the Query</param>
        /// <param name="outputFormat">The <see cref="ReportFormat"/> for the Output of the Report.</param>
        /// <param name="emailRecipients">The List of strings for the Email Addresses of Recipients for the Report</param>
        /// <param name="saveToDisk">Value to indicate whether to save and keep the Report on Disk</param>
        public StoredProcedureReport(string procedureName,
            List<SqlParameter> parameters,
            ReportFormat outputFormat,
            List<string> emailRecipients,
            bool saveToDisk = false)
        {
            ProcedureName = procedureName;
            Parameters = parameters;
            OutputFormat = outputFormat;
            EmailRecipients = emailRecipients;
            this.SaveToLocalDisk = saveToDisk;
        }

        /// <summary>
        /// Creates a new Instance of the <see cref="StoredProcedureReport"/> Class
        /// </summary>
        /// <param name="procedureName">The Name of the Stored Procedure to run</param>
        /// <param name="parameters">The <see cref="List{SqlParameter}"/> to add to the Query</param>
        /// <param name="outputFormat">The <see cref="ReportFormat"/> for the Output of the Report.</param>
        /// <param name="emailRecipients">The List of strings for the Email Addresses of Recipients for the Report</param>
        public StoredProcedureReport(string procedureName, 
            List<SqlParameter> parameters, 
            ReportFormat outputFormat, 
            List<string> emailRecipients)
        {
            ProcedureName = procedureName;
            Parameters = parameters;
            OutputFormat = outputFormat;
            EmailRecipients = emailRecipients;
        }

        /// <summary>
        /// Creates a new Instance of the <see cref="StoredProcedureReport"/> Class
        /// </summary>
        /// <param name="procedureName">The Name of the Stored Procedure to run</param>
        /// <param name="outputFormat">The <see cref="ReportFormat"/> for the Output of the Report</param>
        /// <param name="saveToDisk">Value to indicate whether to save to local Disk or not</param>
        public StoredProcedureReport(string procedureName, ReportFormat outputFormat, bool saveToDisk = false)
        {
            this.ProcedureName = procedureName;
            this.OutputFormat = outputFormat;
            this.Parameters = new List<SqlParameter>();
            this.EmailRecipients = new List<string>();
            this.SaveToLocalDisk = saveToDisk;
        }
    }
}