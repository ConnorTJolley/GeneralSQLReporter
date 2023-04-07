namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Mail;

    public static class GeneralExtensions
    {
        /// <summary>
        /// Converts a <see cref="List{MailAddress}"/>' into a <see cref="List{string}"/>
        /// </summary>
        /// <param name="addresses">The parameter to Extend and convert</param>
        /// <returns>The Resulting <see cref="List{string}"/></returns>
        public static List<string> ToListOfAddresses(this List<MailAddress> addresses) =>
            (from a in addresses 
             where !string.IsNullOrWhiteSpace(a.Address.Trim()) 
             select a.Address)
            .ToList();

        /// <summary>
        /// Convers a <see cref="List{string}"/> of Email Addresses into a <see cref="List{MailAddress}"/>
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
        /// Convers a <see cref="List{string}"/> of Email Addresses into a delimeted list
        /// </summary>
        /// <param name="addresses">The <see cref="List{string}"/> to extend and convert</param>
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
    }
}