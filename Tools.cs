using System;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices.AccountManagement;

namespace useful_tools_dotnet
{
    /// <summary>
    /// A bag of unrelated tools that's useful in small .net app development. Many of which are stolen from StackOverflow.
    /// </summary>
    public class Tools
    {
        /// <summary>
        /// Supports ToSize function.
        /// </summary>
        public enum SizeUnits
        {
            Byte, KB, MB, GB, TB, PB, EB, ZB, YB
        }

        /// <summary>
        /// Converts given value to unit of your choice.
        /// </summary>
        /// <param name="value">Size in bytes.</param>
        /// <param name="unit">Unit for conversion.</param>
        /// <returns>Target unit to two decimal places.</returns>
        public static string ToSize(Int64 value, SizeUnits unit)
        {
            return (value / (double)Math.Pow(1024, (Int64)unit)).ToString("0.00") + unit.ToString();
        }

        /// <summary>
        /// Gets raw tabular data from database.
        /// </summary>
        /// <param name="connectionString">Database to connect.</param>
        /// <param name="query">Query intended for server.</param>
        /// <param name="param">Any SqlParameters needed.</param>
        public static DataTable GetRawData(string connectionString, string query, SqlParameter param = null)
        {
            DataTable table = new DataTable();
            using (var cmd = new SqlCommand(query, new SqlConnection(connectionString)))
            {
                if (param != null) cmd.Parameters.Add(param);
                cmd.Connection.Open();
                table.Load(cmd.ExecuteReader());
            }
            return table;
        }

        /// <summary>
        /// Retrieves specified resource as string.
        /// </summary>
        public static string GetResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            string result;
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                result = reader.ReadToEnd();
            }
            return result;
        }

        /// <summary>
        /// Writes the given text to temporary file path and opens using default handler.
        /// </summary>
        public static void WriteToTempAndOpen(string fileName, string fileTextContents)
        {
            var path = Path.GetTempPath() + @"\" + fileName;
            File.WriteAllText(path, fileTextContents);
            Process.Start(path);
        }

        /// <summary>
        /// Sends a email notification. Set IsSuccessful to true if successful. Msg should contain email body.
        /// </summary>
        public static void SendNotificationEmail(bool IsSuccessful, string msg, string title, string emailServer, string toAddy)
        {
            SmtpClient smtpClient = new SmtpClient();
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress("no-reply@app.dev");

            smtpClient.Host = emailServer;
            smtpClient.UseDefaultCredentials = false;

            message.From = fromAddress;

            string subject = string.Empty;
            if (IsSuccessful) subject = "Success - ";
            else subject = "Failed - ";
            message.Subject = subject + title;

            message.IsBodyHtml = true;
            message.Body = msg;
            message.To.Add(toAddy);

            smtpClient.Send(message);
        }

        /// <summary>
        /// Sends a email notification. Msg should contain email body.
        /// </summary>
        public static void SendNotificationEmail(string msg, string title, string emailServer, string toEmail, string from = "no-reply@app.dev")
        {
            SmtpClient smtpClient = new SmtpClient();
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress(from);

            smtpClient.Host = emailServer;
            smtpClient.UseDefaultCredentials = false;

            message.From = fromAddress;

            string subject = string.Empty;
            message.Subject = subject + title;

            message.IsBodyHtml = true;
            message.Body = msg;
            message.To.Add(toEmail);

            smtpClient.Send(message);
        }
        
        /// <summary>
        /// Checks if current user is a member of the specified group.
        /// </summary>
        /// <param name="groupName">Active Directory secure group to be checked.</param>
        public static bool IsMemberOfAdSecurityGroup(string groupName, string domain)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain);
            UserPrincipal user = UserPrincipal.Current;
            using(PrincipalSearchResult<Principal> groups = user.GetAuthorizationGroups())
            {
                return groups.OfType<GroupPrincipal>().Any(g => g.Name.Equals(groupName, StringComparison.OrdinalIgnoreCase));
            }
        }
    }
}