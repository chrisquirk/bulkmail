using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;

namespace bulkmail
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.Error.WriteLine("ERROR need two args: template maillist.tsv");
                Environment.Exit(1);
            }

            var template = File.ReadAllText(args[0]);
            var maillist = ReadTsv(args[1]).ToList();

            var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            NetworkCredential nc;
            Credential.GetCredentialsVistaAndUp("office365", out nc);
            service.Credentials = nc;
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.Url = new Uri("https://outlook.office365.com/ews/exchange.asmx");

            Regex r = new Regex(@"^\s*([^:]+):\s*(.*)$");
            foreach (var item in maillist)
            {
                var mail = new EmailMessage(service);

                var s = template;
                foreach (var kvp in item)
                    s = s.Replace("[[" + kvp.Key + "]]", kvp.Value);

                bool foundEmpty = false;
                StringBuilder body = new StringBuilder();
                foreach (var line in s.Split('\n'))
                {
                    if (foundEmpty) { body.AppendLine(line); continue; }
                    if (string.IsNullOrWhiteSpace(line)) { foundEmpty = true; continue; }
                    var m = r.Match(line);
                    if (!m.Success) throw new Exception("bad header field " + line);
                    switch (m.Groups[1].Value.ToLowerInvariant())
                    {
                        case "subject": mail.Subject = m.Groups[2].Value.Trim(); break;
                        case "to": mail.ToRecipients.Add(m.Groups[2].Value.Trim()); break;
                        case "cc": mail.CcRecipients.Add(m.Groups[2].Value.Trim()); break;
                        case "bcc": mail.BccRecipients.Add(m.Groups[2].Value.Trim()); break;
                        case "attach": mail.Attachments.AddFileAttachment(m.Groups[2].Value.Trim()); break;

                        default:
                            throw new Exception("unknown field in template: " + m.Groups[1].Value);
                    }
                }
                mail.Body = body.ToString();

                mail.SendAndSaveCopy();
            }
        }

        public static IEnumerable<Dictionary<string, string>> ReadTsv(string filename)
        {
            var header = File.ReadLines(filename).First().Split('\t').ToArray();
            foreach (var l in File.ReadLines(filename).Skip(1))
            {
                var d = new Dictionary<string, string>();
                var v = l.Split('\t').ToArray();
                for (int i = 0; i < header.Length; ++i)
                    d[header[i]] = v[i];
                yield return d;
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            Console.WriteLine("validating " + redirectionUrl);
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
