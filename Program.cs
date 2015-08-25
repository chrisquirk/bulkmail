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
        [Help("send bulk emails using office365 or an exchange server")]
        class Options
        {
            [PositionalArg(0), Required, Help("template for email; starting with headers, then HTML email body")]
            public string Template { get; set; }
            [PositionalArg(1), Required, Help("list of recipients as a tab-separated value file; first line is a header")]
            public string Maillist { get; set; }

            [NamedArg(LongForm="preview", ShortForm='p'), Help("if true, then just write html emails but don't send.")]
            public bool Preview { get; set; }
            [NamedArg(LongForm="traceexch", ShortForm='t'), Help("if true, debug connection with exchange")]
            public bool TraceExchange { get; set; }
            [NamedArg(LongForm="previewdir", ShortForm='d'), DefaultValue("preview"), Help("directory where previews should be written")]
            public string PreviewDir { get; set; }
            [NamedArg(LongForm = "server", ShortForm = 's'), DefaultValue("https://outlook.office365.com/ews/exchange.asmx"), Help("exchange server")]
            public string ExchangeServer { get; set; }
        }

        static void Main(string[] args)
        {
            try
            {
                var o = CommandLineParser.ParseCommandLine<Options>(args);

                var template = File.ReadAllText(o.Template);
                var maillist = ReadTsv(o.Maillist).ToList();

                var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                NetworkCredential nc;
                Credential.GetCredentialsVistaAndUp("office365", out nc);
                service.Credentials = nc;
                if (o.TraceExchange)
                {
                    service.TraceEnabled = true;
                    service.TraceFlags = TraceFlags.All;
                }
                service.Url = new Uri(o.ExchangeServer);

                if (o.Preview && !Directory.Exists(o.PreviewDir))
                    Directory.CreateDirectory(o.PreviewDir);

                Regex r = new Regex(@"^\s*([^:]+):\s*(.*)$");
                int i = 0;
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

                    if (o.Preview)
                    {
                        using (var sw = new StreamWriter(Path.Combine(o.PreviewDir, string.Format("preview_{0}.html", ++i))))
                        {
                            sw.WriteLine("<pre>");
                            sw.WriteLine("Subject: " + mail.Subject);
                            foreach (var email in mail.ToRecipients)
                                sw.WriteLine("To: " + email.ToString());
                            foreach (var email in mail.CcRecipients)
                                sw.WriteLine("Cc: " + email.ToString());
                            foreach (var email in mail.BccRecipients)
                                sw.WriteLine("Bcc: " + email.ToString());
                            foreach (var attach in mail.Attachments)
                                sw.WriteLine("Attach: " + attach.Name);
                            sw.WriteLine("</pre>");
                            sw.WriteLine("<hr>");
                            sw.WriteLine(mail.Body);
                        }
                        continue;
                    }


                    mail.SendAndSaveCopy();
                }
            }
            catch (CommandLineParseError err)
            {
                Console.Error.WriteLine(err.Message);
                Console.Error.WriteLine(err.Usage);
                Environment.Exit(1);
            }
            catch (Exception exn)
            {
                Console.Error.WriteLine(exn.ToString());
                Environment.Exit(1);
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
