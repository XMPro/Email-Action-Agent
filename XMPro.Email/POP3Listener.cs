using MailKit.Net.Pop3;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using XMIoT.Framework;
using XMIoT.Framework.Settings;

namespace XMPro.Email
{
    public class POP3Listener : IAgent, IPollingAgent
    {
        private Configuration config;
        private int mailCount;

        public long UniqueId { get; set; }

        public bool DisableSSLValidation
        {
            get
            {
                var temp = false;
                bool.TryParse(this.config["DisableSSLValidation"], out temp);
                return temp;
            }
        }

        public string Host => this.config["Host"];
        public int Port => int.TryParse(this.config["Port"], out var port) ? port : 995;
        public bool UseSSL => bool.TryParse(this.config["UseSSL"], out var useSSl) ? useSSl : true;

        public string Password
        {
            get
            {
                var request = new OnDecryptRequestArgs(this.config["Password"]);
                this.OnDecryptRequest?.Invoke(this, request);
                return request.DecryptedValue;
            }
        }

        public string Username => this.config["Username"];
        public bool DeleteOnRead => bool.TryParse(this.config["DeleteOnRead"], out var deleteOnRead) ? deleteOnRead : false;

        public event EventHandler<OnPublishArgs> OnPublish;

        public event EventHandler<OnDecryptRequestArgs> OnDecryptRequest;

        public void Create(Configuration configuration)
        {
            this.config = configuration;
        }

        public void Destroy()
        {
        }

        public string GetConfigurationTemplate(string template, IDictionary<string, string> parameters)
        {
            var settingsObj = Settings.Parse(template);
            new Populator(parameters).Populate(settingsObj);

            return settingsObj.ToString();
        }

        public IEnumerable<XMIoT.Framework.Attribute> GetOutputAttributes(string endpoint, IDictionary<string, string> parameters)
        {
            yield return new XMIoT.Framework.Attribute("MessageId", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("ResentMessageId", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("Date", XMIoT.Framework.Settings.Enums.Types.DateTime);
            yield return new XMIoT.Framework.Attribute("ResentDate", XMIoT.Framework.Settings.Enums.Types.DateTime);
            yield return new XMIoT.Framework.Attribute("Importance", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("Priority", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("Sender", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("ResentSender", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("From", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("ResentFrom", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("ReplyTo", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("ResentReplyTo", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("To", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("Cc", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("ResentCc", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("Bcc", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("ResentBcc", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("InReplyTo", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("Subject", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("TextBody", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("HtmlBody", XMIoT.Framework.Settings.Enums.Types.String);
            yield return new XMIoT.Framework.Attribute("Attachments", XMIoT.Framework.Settings.Enums.Types.String);
        }

        public void Poll()
        {
            using (var client = new Pop3Client())
            {
                this.initialize(client);
                var deleted = 0;
                for (; this.mailCount < client.Count; this.mailCount++)
                {
                    try
                    {
                        var message = client.GetMessage(this.mailCount);
                        this.OnPublish?.Invoke(this, new OnPublishArgs(this.composeOutput(message)));
                        if (this.DeleteOnRead)
                        {
                            client.DeleteMessage(this.mailCount);
                            deleted++;
                        }
                    }
                    catch (Exception ex)
                    {
#warning This should be logged properly.
                        Console.WriteLine($"{DateTime.UtcNow}|ERROR|{nameof(POP3Listener)}|{ex.ToString()}");
                    }
                }

                client.Disconnect(true);
                //sucessfully deleted
                if (this.DeleteOnRead)
                    this.mailCount -= deleted;
            }
        }

        public void Start()
        {
            using (var client = new Pop3Client())
            {
                this.initialize(client);

                this.mailCount = client.Count;

                client.Disconnect(true);
            }
        }

        public string[] Validate(IDictionary<string, string> parameters)
        {
            int i = 1;
            var errors = new List<string>();
            this.config = new Configuration() { Parameters = parameters };

            if (String.IsNullOrWhiteSpace(this.Host))
                errors.Add($"Error {i++}: Host address is not specified.");

            if (String.IsNullOrWhiteSpace(this.Username))
                errors.Add($"Error {i++}: Username is not specified.");

            return errors.ToArray();
        }

        private void initialize(Pop3Client client)
        {
            if (this.DisableSSLValidation)
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

            client.Connect(this.Host, this.Port, this.UseSSL);

            client.Authenticate(this.Username, this.Password);
        }

        private Array composeOutput(params MimeMessage[] messages)
        {
            var events = new List<IDictionary<string, object>>();
            foreach (var message in messages)
            {
                var rtr = new Dictionary<string, object>();
                rtr.Add("MessageId", message.MessageId);
                rtr.Add("ResentMessageId", message.ResentMessageId);
                rtr.Add("Date", message.Date.DateTime);
                rtr.Add("ResentDate", message.ResentDate.DateTime);
                rtr.Add("Importance", message.Importance.ToString());
                rtr.Add("Priority", message.Priority.ToString());
                rtr.Add("Sender", message.Sender?.ToString());
                rtr.Add("ResentSender", message.ResentSender?.ToString());
                rtr.Add("From", String.Join(";", message.From));
                rtr.Add("ResentFrom", String.Join(";", message.ResentFrom));
                rtr.Add("ReplyTo", String.Join(";", message.ReplyTo));
                rtr.Add("ResentReplyTo", String.Join(";", message.ResentReplyTo));
                rtr.Add("To", String.Join(";", message.To));
                rtr.Add("Cc", String.Join(";", message.Cc));
                rtr.Add("ResentCc", String.Join(";", message.ResentCc));
                rtr.Add("Bcc", String.Join(";", message.Bcc));
                rtr.Add("ResentBcc", String.Join(";", message.ResentBcc));
                rtr.Add("InReplyTo", message.InReplyTo);
                rtr.Add("Subject", message.Subject);
                rtr.Add("TextBody", message.TextBody);
                rtr.Add("HtmlBody", message.HtmlBody);

                var attachments = new List<string>();
                foreach (MimeEntity attachment in message.Attachments)
                {
                    var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                    var dirPath = this.getRandomDirName();
                    Directory.CreateDirectory(dirPath);
                    var path = Path.Combine(dirPath, fileName);
                    using (var stream = System.IO.File.Create(Path.Combine(dirPath, fileName)))
                    {
                        if (attachment is MessagePart)
                        {
                            var rfc822 = (MessagePart)attachment;
                            rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment;
                            part.Content.DecodeTo(stream);
                        }
                    }
                    attachments.Add(path);
                }
                rtr.Add("Attachments", String.Join("|", attachments));

                events.Add(rtr);
            }
            return events.ToArray();
        }

        private string getRandomDirName()
        {
            var tempDir = System.IO.Path.GetTempPath();

            string fullPath;
            do
            {
                var randomName = System.IO.Path.GetRandomFileName();
                fullPath = System.IO.Path.Combine(tempDir, randomName);
            }
            while (Directory.Exists(fullPath));

            return fullPath;
        }
    }
}