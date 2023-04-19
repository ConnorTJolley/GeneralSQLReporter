namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Linq;
    using System.Net.Mail;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Extension Methods for General Uses / Classes
    /// </summary>
    /// <remarks>
    /// Extension Methods which aren't particularly used for any specific Class, like <see cref="SqlReportExtensions"/>
    /// </remarks>
    public static class GeneralExtensions
    {
        /// <summary>
        /// Handles Converting a <see cref="SqlParameter"/> into it's SQL Declarative value
        /// </summary>
        /// <param name="parameter">The <see cref="SqlParameter"/> to Convert</param>
        /// <returns>The Resulting SQL Declarative Value, e.g nvarchar 'VALUE', int 0</returns>
        internal static string ToQueryString(this SqlParameter parameter)
        {
            switch (parameter.DbType)
            {
                case DbType.String:
                case DbType.Binary:
                case DbType.Date:
                case DbType.Time:
                case DbType.DateTime2:
                case DbType.DateTimeOffset:
                case DbType.Xml:
                case DbType.Guid:
                    return $"'{parameter.Value}'";

                default:
                    return parameter.Value.ToString();
            }
        }

        /// <summary>
        /// Handles Converting the <see cref="List{SqlParameter}"/> into SQL Variables / Params
        /// </summary>
        /// <param name="parameters">The <see cref="List{SqlParameter}"/> to Convert.</param>
        /// <returns>The Resulting SQL Variable Declaration and Assignment.</returns>
        public static string ToSqlParams(this List<SqlParameter> parameters)
        {
            var result = $"-- Parameters{Environment.NewLine}DECLARE ";

            if (!parameters.Any())
            {
                return $"-- No Parameters Used.{Environment.NewLine}";
            }

            parameters.ForEach(param =>
            {
                var upperType = param.SqlDbType.ToString().ToUpper();
                result += $"{param.ParameterName} {upperType}, ";
            });

            //// Remove trailing comma
            var trimmed = result.Trim();
            result = trimmed.Trim().Remove(trimmed.Length - 1);
            result += $";{Environment.NewLine}";

            parameters.ForEach(param =>
            {
                var typeValue = param.ToQueryString();
                result += $"SET {param.ParameterName} = {typeValue};{Environment.NewLine}";
            });

            result += $"-- End of Parameters{Environment.NewLine + Environment.NewLine}";
            return result;
        }

        /// <summary>
        /// Handles the Adding of <see cref="SqlParameter"/>s to the <see cref="SqlCommand"/> if the <see cref="GenericReport"/> has them
        /// </summary>
        /// <param name="command">The <see cref="SqlCommand"/> to extend and add the <see cref="SqlParameter"/>'s to</param>
        /// <param name="report">The <see cref="GenericReport"/> that is being ran</param>
        public static void AddParamsIfRequired(this SqlCommand command, GenericReport report)
        {
            if (report.Parameters.Any())
            {
                report.Parameters.ForEach(param =>
                {
                    command.Parameters.Add(param);
                });
            }
        }

        /// <summary>
        /// Converts a <see cref="List{MailAddress}"/>' into a List of strings
        /// </summary>
        /// <param name="addresses">The parameter to Extend and convert</param>
        /// <returns>The Resulting List of strings</returns>
        public static List<string> ToListOfAddresses(this List<MailAddress> addresses) =>
            (from a in addresses 
             where !string.IsNullOrWhiteSpace(a.Address.Trim()) 
             select a.Address)
            .ToList();

        /// <summary>
        /// Convers a List of strings of Email Addresses into a <see cref="List{MailAddress}"/>
        /// </summary>
        /// <param name="addresses">The parameter to Extend and convert</param>
        /// <returns>The Resulting <see cref="List{MailAddress}"/></returns>
        public static List<MailAddress> ToListOfMailAddresses(this List<string> addresses)
        {
            try
            {
                var mailAddresses = new List<MailAddress>();
                var vals = addresses.Where(a => !string.IsNullOrWhiteSpace(a.Trim())).ToList();
                vals.ForEach(v =>
                {
                    mailAddresses.Add(new MailAddress(v));
                });

                return mailAddresses;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception when Converting List of Addresses into MailAddresses. " +
                    $"Ex: {ex.Message}");
                throw ex;
            }
        }

        /// <summary>
        /// Convers a List of strings of Email Addresses into a delimeted list
        /// </summary>
        /// <param name="addresses">The List of strings to extend and convert</param>
        /// <param name="delimeter">The Delimeter to separate the values with, defaults to a comma</param>
        /// <returns>The results delimeted string</returns>
        public static string ToFullEmailString(this List<string> addresses, char delimeter = ',')
        {
            var fullString = string.Empty;
            addresses?.ForEach(rec =>
            {
                fullString += $"{rec.Trim()}{delimeter} ";
            });

            //// Remove ending delimeter and trailing space.
            fullString = fullString.Trim().Remove(fullString.Length - 1, 1);

            return fullString;
        }

        /// <summary>
        /// Formats a <see cref="List{SqlParameter}"/> into a Formatted String containing Information about each parameter and their value / type
        /// </summary>
        /// <param name="parameters">The <see cref="List{SqlParameter}"/> To Extend and Format</param>
        /// <returns>The Formatted String Result</returns>
        public static string ToFormattedString(this List<SqlParameter> parameters)
        {
            var msg = string.Empty;

            if (parameters.Any())
            {
                msg += $"{parameters.Count} Parameters:{Environment.NewLine}" +
                    "-----" +
                    $"{Environment.NewLine}";

                parameters.ForEach(p =>
                {
                    msg += $"Name: {p.ParameterName}. Type: {p.DbType}. Value: {p.Value}{Environment.NewLine}";
                });

                msg += "-----";
            }

            return msg;
        }
    }
}