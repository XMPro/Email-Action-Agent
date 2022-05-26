# Email Action Agent

## Prerequisites
- Visual Studio
- [XMPro IoT Framework NuGet package](https://www.nuget.org/packages/XMPro.IOT.Framework/)
- Please see the [Manage Agents](https://documentation.xmpro.com/how-tos/manage-agents) guide for a better understanding of how the Agent Framework works

## Description
The Email action agent allows an e-mail to be sent in the stream at any point in the data flow.

## How the code works
All settings referred to in the code need to correspond with the settings defined in the template that has been created for the agent using the XMPro Package Manager. Refer to the [XMPro Package Manager](https://documentation.xmpro.com/agent/packaging-agents/) guide for instructions on how to define the settings in the template and package the agent after building the code. 

After packaging the agent, you can upload it to the XMPro Data Stream Designer and start using it.

### Settings
When a user needs to use the *Email Action Agent*, they need to provide a number of settings. To get the parent outputs, use the following code: 

```csharp
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
```

The user needs to specify an SMTP server and port that can be used.

```csharp
private string SMTPServer { get { return this.config["SMTPServer"]; } }
private int SMTPPort
{
    get
    {
        string port = this.config["SMTPPort"];
        return string.IsNullOrWhiteSpace(port) ? 25 : int.Parse(port);
    }
}
```

If the user prefers to use default credentials, the *Use Default Credentials* check box needs to be selected. If not, he/ she needs to provide a username and password that can be used. 

```csharp
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
```

Next, the user needs to specify an address from which the e-mail will be sent as well as an address to which the e-mail needs to be sent. To get these values, use the code below.

```csharp
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
```

Get the subject and body of the e-mail:

```csharp
private string Subject { get { return this.config["Subject"]; } }
private string Body { get { return this.config["Body"]; } }
```

### Configuration
In the *GetConfigurationTemplate* method, parse the JSON representation of the settings into the Settings object.

```csharp
var settingsObj = Settings.Parse(template);
new Populator(parameters).Populate(settingsObj);
```

Create controls for and set the values for the following items:
* *Use Default Credentials* check box
* *Username* text box
* *Password* text box
* *Subject* text box
* *Body* HTML editor
* *From* token box
* *To* token box
* *Place holders* grid

```csharp
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
```

If the placeholders grid is not empty, create a drop down that will contain the mapping options. These options will come from this agent's parent agent and can be mapped to any place holder added in the *Body* field.

```csharp
DropDown Mapping = PlaceHoldersGrid.Columns.First(s => s.Key == "Mapping") as DropDown;
Mapping.Options = this.ParentOutputs.Select(l => new Option() { DisplayMemeber = l.Name, ValueMemeber = l.Name }).ToList();
```

Get all the template fields.

```csharp
var fields = new List<String>();
fields = GetTemplateFields(Subject.Value, Body.Value).Keys.ToList();
```

### Validate
When validating the stream, an error needs to be shown if there are rows added to the placeholders grid that are not mapped. Thus, if there are placeholders in the e-mail body that are not mapped to values.

```csharp
List<string> errors = new List<string>();
this.config = new Configuration() { Parameters = parameters };
var PlaceHoldersGrid = new Grid();
PlaceHoldersGrid.Value = this.config["PlaceHoldersGrid"];

foreach (var row in PlaceHoldersGrid.Rows)
{
    if (String.IsNullOrEmpty(row["Mapping"].ToString()))
        errors.Add($"Mapping not defined for place holder {row["PlaceHolder"].ToString()}");
}
```

### Create
Set the config variable to the configuration received in the *Create* method. 

```csharp
this.config = configuration;
```

### Start
In the *Start* method, create a new instance of the *SmtpClient* class. Specify the event handler that should be called when the *SmtpClient.SendCompleted* event is raised. This event will be raised after an asynchronous e-mail has been sent. Set the SMTP host and port to the values the user specified and set *Smpt.EnableSsl* to true. 

```csharp
smtp = new SmtpClient();
smtp.SendCompleted += Smtp_SendCompleted;
smtp.Host = SMTPServer;
smtp.Port = SMTPPort;
smtp.EnableSsl = true;
```

Set *Smpt.UseDefaultCredentials* to the value the user specified. If the user did not choose to use default credentials, set the credentials to his/her username as specified on the configuration page. 

```csharp
smtp.UseDefaultCredentials = UseDefaultCredentials;
  if (smtp.UseDefaultCredentials == false)
      smtp.Credentials = new NetworkCredential(this.UserName, this.Password);
```

Add an event handler to your code to handle the *SendCompleted* event.

```csharp
private async void Smtp_SendCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
{
    if (e.Error != null)
    {
        //await this.Reporter.Error(this.Properties.Name, "Failed to send Email.", e.Error).ConfigureAwait(false);
    }
}
```

### Destroy
There is no need to do anything in the *Destroy* method.

### Publishing Events
Because this agent will be receiving events, the *Receive* method needs to be implemented. Create a new Grid and set its value to the PlaceHolders grid.

```csharp
var grid = new Grid();
grid.Value = this.config["PlaceHolders"];
```

For each of the events received, start by getting the event, convert it to a dictionary and then create a new message. Next, set the subject and body of the e-mail. Lastly, send the e-mail.

```csharp
foreach (JObject _event in events)
{
    var rtr = this.ToDictionary(_event);
    msg = GetNewMessage(_event);
    msg.Subject = ParseText(Subject, grid, _event);
    msg.Body = ParseText(Body, grid, _event);
    smtp.Send(msg);
}
```

Publish the events by invoking the OnPublish event.

```csharp
this.OnPublish?.Invoke(this, new OnPublishArgs(events));
```
### Decrypting Values
Since this agent needs secure settings (Password), the values will automatically be encrypted. Use the following code to decrypt the values.

```csharp
private string decrypt(string value)
{
    var request = new OnDecryptRequestArgs(value);
    this.OnDecryptRequest?.Invoke(this, request);
    return request.DecryptedValue;
}
```

### Helper Methods
### Get New Message
The *Email Action Agent* uses several helper methods. The first of these methods creates and returns a new e-mail message. The newly created message is of type *MailMessage*, which represents an e-mail message that can be sent using the *SmptClient* class. In this method, start by creating a new instance of the *MailMessage* class and set the *IsBodyHtml* property to *true*. Next, set the e-mail address from which the e-mail should be sent. Make sure that the *To* field of the message object is cleared and add each e-mail address that the user specified to the list of addresses the e-mail needs to be sent to.

```csharp
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
```

### Get Template Fields
The *GetTemplateFields* method accepts a subject and body of an e-mail, both of type *String*. This method returns a dictionary containing the template fields. To get the template fields, a regular expression is used to determine if the subject and body of an e-mail contains words that are wrapped by certain characters, marking them as placeholders, for example, *{val1}*.

```csharp
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
```

### Parse Text
A helper method is needed to replace the placeholders that have been added in the e-mail subject or body with the values they are mapped to, for example "{UserFullName}" would be replaced with a value such as "Keith Smith".

```csharp
private string ParseText(string subject, Grid grid, JObject substitutions)
{
    if (grid.Rows.ToList().Count == 0)
        return subject;

    foreach (var dr in grid.Rows)
        subject = subject.Replace("{" + dr["PlaceHolder"].ToString().Trim() + "}", substitutions[dr["Mapping"].ToString()].ToString() ?? "");

    return subject;
}
```

### To Dictionary
The last helper method we need simply converts a given object to a dictionary. It retrieves the properties of the object and, for each property, gets the value of the property and adds it to a dictionary. This method makes use of the *PropertyDescriptorCollection* class, which represents a collection of *PropertyDescriptor* objects and is part of the *System.ComponentModel* namespace.

```csharp
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
```