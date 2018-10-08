using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using MsgReader.Outlook;
using Newtonsoft.Json.Linq;
using XMIoT.Framework;
using XMIoT.Framework.Settings;

namespace XMPro.Email
{
    public class ActionAgent : IAgent, IReceivingAgent
    {
        #region Global Variables

        private Configuration config;
        private List<XMIoT.Framework.Attribute> parentOutputs;
        private SmtpClient smtp;
        private MailMessage msg;
        private Storage.Message msgTemplate = null;
        private DataTable placeholdersTable = new DataTable();

        private List<XMIoT.Framework.Attribute> ParentOutputs {
            get
            {
                if (this.parentOutputs == null)
                {
                    var args = new OnRequestParentOutputAttributesArgs(this.UniqueId, "Input");
                    this.OnRequestParentOutputAttributes.Invoke(this, args);
                    this.parentOutputs =  args.ParentOutputs.ToList();
                }
                return this.parentOutputs;
            }
        }
        private string SMTPServer { get { return this.config["SMTPServer"]; } }
        private int SMTPPort
        {
            get
            {
                string port = this.config["SMTPPort"];
                return string.IsNullOrWhiteSpace(port) ? 25 : int.Parse(port);
            }
        }
        private bool UseDefaultCredentials
        {
            get
            {
                bool UseDefaultCredentials = false;
                bool.TryParse(this.config["UseDefaultCredentials"], out UseDefaultCredentials);
                return UseDefaultCredentials;
            }
        }
        private string UserName { get { return this.config["UserName"]; } }
        private string Password { get { return decrypt(this.config["Password"]); } }
        private string From
        {
            get
            {
                var from = this.config["From"];
                return from == null ? string.Empty : from.Split(";".ToCharArray())[0];
            }
        }
        private string[] To
        {
            get
            {
                var to = this.config["To"];
                return to == null ? new string[0] : to.Split(";".ToCharArray());
            }
        }
        private string Subject { get { return this.config["Subject"]; } }

        private string Body { get { return this.config["Body"]; } }

        #endregion

        #region Agent Implementation

        public long UniqueId { get; set; }
        public event EventHandler<OnPublishArgs> OnPublish;
        public event EventHandler<OnDecryptRequestArgs> OnDecryptRequest;
        public event EventHandler<OnRequestParentOutputAttributesArgs> OnRequestParentOutputAttributes;

        private string decrypt(string value)
        {
            var request = new OnDecryptRequestArgs(value);
            this.OnDecryptRequest?.Invoke(this, request);
            return request.DecryptedValue;
        }

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

            CheckBox UseDefaultCredentials = settingsObj.Find("UseDefaultCredentials") as CheckBox;
            TextBox UserName = settingsObj.Find("UserName") as TextBox;
            TextBox Password = settingsObj.Find("Password") as TextBox;
            UserName.Visible = Password.Visible = UseDefaultCredentials.Value == false;
            TextBox Subject = settingsObj.Find("Subject") as TextBox;
            HtmlEditor Body = settingsObj.Find("Body") as HtmlEditor;
            Body.Visible = Subject.Visible = true;
            DropDown From = settingsObj.Find("From") as DropDown;
            From.Options = ParentOutputs.Select(i => new Option() { DisplayMemeber = i.Name, ValueMemeber = i.Name }).ToList();
            DropDown To = settingsObj.Find("To") as DropDown;
            To.Options = ParentOutputs.Select(i => new Option() { DisplayMemeber = i.Name, ValueMemeber = i.Name }).ToList();
            var PlaceHoldersGrid = settingsObj.Find("PlaceHolders") as Grid;
            PlaceHoldersGrid.DisableDelete = true;
            PlaceHoldersGrid.DisableInsert = true;

            if (PlaceHoldersGrid != null)
            {
                DropDown Mapping = PlaceHoldersGrid.Columns.First(s => s.Key == "Mapping") as DropDown;
                Mapping.Options = this.ParentOutputs.Select(l => new Option() { DisplayMemeber = l.Name, ValueMemeber = l.Name }).ToList();

                var fields = new List<String>();
                fields = GetTemplateFields(Subject.Value, Body.Value).Keys.ToList();
                
                var newRows = new JArray();
                var rows = PlaceHoldersGrid.Rows?.ToList() ?? new List<IDictionary<string, object>>();
                foreach (var row in rows)
                {
                    if (fields.Contains(row["PlaceHolder"].ToString()) == true)
                        newRows.Add(JObject.FromObject(row));
                }

                foreach (var field in fields)
                {
                    bool foundMatch = false;
                    foreach (JObject row in newRows.Children<JObject>())
                    {
                        var d = ToDictionary(row);
                        if (d.ContainsKey("PlaceHolder") && d["PlaceHolder"].ToString() == field)
                        {
                            foundMatch = true;
                            break;
                        }
                    }

                    if (!foundMatch)
                    {
                        var newRow = new JObject();
                        newRow.Add("PlaceHolder", field);
                        newRow.Add("Mapping", "");
                        var idx = fields.IndexOf(field) < newRows.Count ? fields.IndexOf(field) : newRows.Count;
                        newRows.Insert(idx, newRow);
                    }
                }
                PlaceHoldersGrid.Value = newRows.ToString();
            }

            return settingsObj.ToString();
        }

        public IEnumerable<XMIoT.Framework.Attribute> GetInputAttributes(string endpoint, IDictionary<string, string> parameters)
        {
            return ParentOutputs;
        }

        public IEnumerable<XMIoT.Framework.Attribute> GetOutputAttributes(string endpoint, IDictionary<string, string> parameters)
        {
            return ParentOutputs;
        }
        
        public void Start()
        {
            smtp = new SmtpClient();
            smtp.SendCompleted += Smtp_SendCompleted;
            smtp.Host = SMTPServer;
            smtp.Port = SMTPPort;
            smtp.EnableSsl = true;

            smtp.UseDefaultCredentials = UseDefaultCredentials;
            if (smtp.UseDefaultCredentials == false)
                smtp.Credentials = new NetworkCredential(this.UserName, this.Password);
        }

        public void Receive(string endpointName, JArray events)
        {
            var grid = new Grid();
            grid.Value = this.config["PlaceHolders"];

            foreach (JObject _event in events)
            {
                var rtr = this.ToDictionary(_event);
                msg = GetNewMessage(_event);
                msg.Subject = ParseText(Subject, grid, _event);
                msg.Body = ParseText(Body, grid, _event);
                smtp.Send(msg);
            }

            this.OnPublish?.Invoke(this, new OnPublishArgs(events));
        }

        public string[] Validate(IDictionary<string, string> parameters)
        {
            List<string> errors = new List<string>();
            this.config = new Configuration() { Parameters = parameters };
            var PlaceHoldersGrid = new Grid();
            PlaceHoldersGrid.Value = this.config["PlaceHoldersGrid"];

            foreach (var row in PlaceHoldersGrid.Rows)
            {
                if (String.IsNullOrEmpty(row["Mapping"].ToString()))
                    errors.Add($"Mapping not defined for place holder {row["PlaceHolder"].ToString()}");
            }

            return errors.ToArray();
        }

        #endregion

        #region Helper methods

        private MailMessage GetNewMessage(JObject record)
        {
            var msg = new MailMessage();
            msg.IsBodyHtml = true;

            if (From.IndexOf("@") < 0)
                msg.From = new MailAddress(record[From].ToString());
            else
                msg.From = new MailAddress(From);

            msg.To.Clear();
            foreach (var to in To)
            {
                if (to.IndexOf("@") < 0)
                    msg.To.Add(new MailAddress(record[to].ToString()));
                else
                    msg.To.Add(new MailAddress(to));
            }

            return msg;
        }

        private async void Smtp_SendCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                //await this.Reporter.Error(this.Properties.Name, "Failed to send Email.", e.Error).ConfigureAwait(false);
            }
        }

        private Dictionary<string, string> GetTemplateFields(String subject, String body)
        {
            Dictionary<string, string> fields = new Dictionary<string, string>();
            Regex rx = new Regex(@"\{(.*?)\}");

            if (!String.IsNullOrEmpty(subject))
            {
                MatchCollection mc = rx.Matches(subject);
                foreach (Match match in mc)
                {
                    string hit = match.Value.Substring(1, match.Value.Length - 2);
                    if (!fields.ContainsKey(hit))
                        fields.Add(hit.Trim(), string.Empty);
                }
            }
            if (!String.IsNullOrEmpty(body))
            {
                MatchCollection mc2 = rx.Matches(body);
                foreach (Match match in mc2)
                {
                    string hit = match.Value.Substring(1, match.Value.Length - 2);
                    if (!fields.ContainsKey(hit))
                        fields.Add(hit.Trim(), string.Empty);
                }
            }

            return fields;
        }

        private string ParseText(string subject, Grid grid, JObject substitutions)
        {
            if (grid.Rows.ToList().Count == 0)
                return subject;

            foreach (var dr in grid.Rows)
                subject = subject.Replace("{" + dr["PlaceHolder"].ToString().Trim() + "}", substitutions[dr["Mapping"].ToString()].ToString() ?? "");

            return subject;
        }

        private IDictionary<string, object> ToDictionary(JObject obj)
        {
            IDictionary<string, object> result = new Dictionary<string, object>();
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(obj);
            foreach (PropertyDescriptor property in properties)
            {
                var val = property.GetValue(obj);
                if (val != null && val is JValue)
                    result.Add(property.Name, (val as JValue)?.Value);
                else
                    result.Add(property.Name, val);
            }
            return result;
        }

        #endregion
    }
}
