using System;
using System.Configuration;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Microsoft.Exchange.WebServices.Data;
using EWSMAPI;

/// <summary>
/// Summary description for WHARTONEWS
/// </summary>
[WebService(Namespace = "https://accts.wharton.upenn.edu/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WhartonEWS : System.Web.Services.WebService {

    private static String ItemUrlProtocol = ConfigurationManager.AppSettings["ItemUrlProtocol"];
    private static String ItemUrlFolder = ConfigurationManager.AppSettings["ItemUrlFolder"];

    public WhartonEWS(){}

    private ExchangeService getService(String UserEmailAddress, String apiKey)
    {
        ExchangeService svc = new ExchangeService(ExchangeVersion.Exchange2010);

        String impDomain = ConfigurationManager.AppSettings["ImpersonationDomain_" + apiKey];
        String impUser = ConfigurationManager.AppSettings["ImpersonationUsername_" + apiKey];
        String impPass = ConfigurationManager.AppSettings["ImpersonationPassword_" + apiKey];
        String impEmail = impUser + '@' + impDomain;
        
        //make sure api key is valid
        if (impDomain + impUser + impPass == "")
        {
            throw new Exception("401 - API Key not valid or not found! (API Key: " + apiKey + ")");
        }

        //login as user capable of impersonation
        svc.Credentials = new WebCredentials(impUser, impPass, impDomain);

        //use autodiscover to find the email server
        //svc.AutodiscoverUrl(impEmail);
        svc.Url = new Uri("https://your.exchange.server.com/path/to/EWS.asmx");

        //impersonate the indicated user
        svc.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, UserEmailAddress);

        //allow up to 30 seconds for Exchange to respond
        svc.Timeout = (1000 * 30); //convert 30 seconds into milliseconds


        if (ConfigurationManager.AppSettings["trace"] == "true")
        {
            svc.TraceListener = new TraceListener(ConfigurationManager.AppSettings["traceDir"]);
            svc.TraceFlags = TraceFlags.All;
            svc.TraceEnabled = true;
        }

        return svc;
    }

    #region Meta

    private string cleanHighAsciiForXML(string orig)
    {
        if (orig == null) { return null; }

        Regex highAsciiPattern = new Regex("[^\x20-\x7F\x09\x0D]");

        string result = highAsciiPattern.Replace(orig, delegate(Match match)
        {
            string v = match.ToString();
            string r = "";
            //loop over every character in the match
            for (int iterator = 0; iterator < v.Length; iterator++)
            {
                int i = Convert.ToInt32(v[iterator]);
                switch (i)
                {
                    //ms word smart double-quotes
                    case (8220):
                        r += "\"";
                        break;
                    case (8221):
                        r += "\"";
                        break;

                    //ms word smart single-quotes
                    case (8216):
                        r += "'";
                        break;
                    case (8217):
                        r += "'";
                        break;

                    //ms word smart ellipses
                    case (8230):
                        r += "...";
                        break;
                    
                    //everything else just gets digit-escaped
                    default:
                        r += ("&#" + i.ToString() + ";");
                        break;
                }
            }
            return r;
        });
        return result.Trim();
    }

    /* This method is not currently available as a webservice, as we are still evaluating the possibility of using it. (doesn't work yet either) */
    //[WebMethod]
    private void GetUserMailboxSize(String UserEmailAddress, String apiKey)
    {
        ExchangeService svc = getService(UserEmailAddress, apiKey);

        Folder root = Folder.Bind(svc, WellKnownFolderName.MsgFolderRoot);

        root.Load(new PropertySet(BasePropertySet.FirstClassProperties));

        FindFoldersResults r = root.FindFolders(new FolderView(10000));


        Console.Write(root.ToString());

        int totalSize = traverseChildFoldersForSize(root);

        Console.Write(totalSize.ToString());
    }

    private int traverseChildFoldersForSize(Folder f)
    {
        int folderSizeSum = 0;
        if (f.ChildFolderCount > 0)
        {
            foreach (Folder c in f.FindFolders(new FolderView(10000)))
            {
                folderSizeSum += traverseChildFoldersForSize(c);
            }
        }

        folderSizeSum += (int)f.ManagedFolderInformation.FolderSize;

        return folderSizeSum;
    }

    private void addAuditLog
    (
        String apiKey,
        String impersonatedUser,
        String apiMethod,
        Boolean successFlag,
        String miscData
    )
    {
        DateTime dtTimestamp = DateTime.Now;
        String strIPAddress = Context.Request.ServerVariables["Remote_ADDR"];

        //default to empty json structure
        if (miscData == "")
        {
            miscData = "{}";
        }

        //run the stored procedure to log activity
        SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["Logger"].ConnectionString);
        cn.Open();
        SqlCommand cmd = new SqlCommand("usp_log_activity", cn);
        cmd.CommandType = CommandType.StoredProcedure;
        cmd.Parameters.Add("@logTimestamp", SqlDbType.DateTime).Value = dtTimestamp;
        cmd.Parameters.Add("@APIKey", SqlDbType.VarChar).Value = apiKey;
        cmd.Parameters.Add("@ImpersonatedUser", SqlDbType.VarChar).Value = impersonatedUser;
        cmd.Parameters.Add("@APIMethod", SqlDbType.VarChar).Value = apiMethod;
        cmd.Parameters.Add("@MethodSuccess", SqlDbType.Bit).Value = successFlag;
        cmd.Parameters.Add("@IPAddress", SqlDbType.VarChar).Value = strIPAddress;
        cmd.Parameters.Add("@MiscData", SqlDbType.VarChar).Value = miscData;
        int rtnCode = cmd.ExecuteNonQuery();
        if (cn != null)
        {
            cn.Close();
        }
    }

    #endregion

    #region Mail

    [WebMethod]
    public WhartonEWSResponse CreateEmail
    (
        String apiKey,
        String emlUserAddress, 
        String emlTORecipients, //comma-delimited list
        String emlCCRecipients, //comma-delimited list
        String emlBCCRecipients, //comma-delimited list
        String strSubject, 
        String strBody, 
        String strImportance
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            //create the message object
            EmailMessage msg = new EmailMessage(service);

            //fill in its properties
            msg.Subject = strSubject;
            msg.Body = strBody;

            if (strImportance.ToLower() == "high")
                msg.Importance = Importance.High;
            else if (strImportance.ToLower() == "low")
                msg.Importance = Importance.Low;
            else if (strImportance.ToLower() == "normal")
                msg.Importance = Importance.Normal;
            else
            {
                rsp.Msg = "Unrecognized importance level, using 'Normal'";
                msg.Importance = Importance.Normal;
            }

            //add TO address(es). TO is required for all emails, so return false if none have been supplied.
            if (emlTORecipients.Length == 0)
            {
                throw new Exception("400 - TO recipients is empty. You must specify at least one TO recipient. Mail not sent.");
            }
            else
            {
                emlTORecipients.Replace(';', ',');
                String[] TOaddresses = emlTORecipients.Split(',');
                foreach (String e in TOaddresses)
                {
                    msg.ToRecipients.Add(e.Trim());
                }
            }
            if (emlCCRecipients.Length > 0)
            {
                emlCCRecipients.Replace(';', ',');
                String[] CCaddresses = emlCCRecipients.Split(',');
                foreach (String e in CCaddresses)
                {
                    msg.CcRecipients.Add(e.Trim());
                }
            }
            if (emlBCCRecipients.Length > 0)
            {
                emlBCCRecipients.Replace(';', ',');
                String[] BCCaddresses = emlBCCRecipients.Split(',');
                foreach (String e in BCCaddresses)
                {
                    msg.BccRecipients.Add(e.Trim());
                }
            }

            //send the email and save a copy to the sent items folder
            msg.SendAndSaveCopy();

        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "CreateEmail", success, miscData);

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse GetEmail
    (
        String apiKey,
        String emlUserAddress, 
        int intNumMessages
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            FindItemsResults<Item> messages = getEmailMessages(apiKey, emlUserAddress, intNumMessages);

            // Create a DataTable. 
            DataTable table = new DataTable("Items");

            #region table column definitions

            DataColumn dcFrom = new DataColumn();
            dcFrom.DataType = System.Type.GetType("System.String");
            dcFrom.AllowDBNull = true;
            dcFrom.Caption = "From";
            dcFrom.ColumnName = "From";
            dcFrom.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcFrom);

            DataColumn dcTo = new DataColumn();
            dcTo.DataType = System.Type.GetType("System.String");
            dcTo.AllowDBNull = true;
            dcTo.Caption = "To";
            dcTo.ColumnName = "To";
            dcTo.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcTo);

            DataColumn dcSubject = new DataColumn();
            dcSubject.DataType = System.Type.GetType("System.String");
            dcSubject.AllowDBNull = true;
            dcSubject.Caption = "Subject";
            dcSubject.ColumnName = "Subject";
            dcSubject.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcSubject);

            DataColumn dcLink = new DataColumn();
            dcLink.DataType = System.Type.GetType("System.String");
            dcLink.AllowDBNull = true;
            dcLink.Caption = "Link";
            dcLink.ColumnName = "Link";
            dcLink.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLink);

            DataColumn dcDateTimeSent = new DataColumn();
            dcDateTimeSent.DataType = System.Type.GetType("System.DateTime");
            dcDateTimeSent.AllowDBNull = true;
            dcDateTimeSent.Caption = "DateTimeSent";
            dcDateTimeSent.ColumnName = "DateTimeSent";

            #endregion

            // Add the column to the table. 
            table.Columns.Add(dcDateTimeSent);

            // Add rows and set values. 
            DataRow row;
            foreach (EmailMessage msg in messages)
            {
                string openItemUrl = ItemUrlProtocol + service.Url.Host + ItemUrlFolder + msg.WebClientReadFormQueryString;

                row = table.NewRow();
                row["From"] = cleanHighAsciiForXML(msg.From.Name);
                row["To"] = cleanHighAsciiForXML(msg.DisplayTo);
                row["Subject"] = cleanHighAsciiForXML(msg.Subject);
                row["Link"] = openItemUrl;
                row["DateTimeSent"] = msg.DateTimeSent;

                // Be sure to add the new row to the 
                // DataRowCollection. 
                table.Rows.Add(row);
            }

            rsp.TableData = table;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetEmail", success, miscData);

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse GetEmailUnreadCount
    (
        String apiKey,
        String emlUserAddress
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            ExchangeService svc = getService(emlUserAddress, apiKey);
            Folder root = Folder.Bind(svc, WellKnownFolderName.Inbox);
            int unread = root.UnreadCount;
            rsp.SimpleData = unread;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetEmailUndreadCount", success, miscData);

        return rsp;
    }

    //used by other email functionality, single point of maintenance for mail lookup functionality.
    private FindItemsResults<Item> getEmailMessages
    (
        String apiKey,
        String emlUserAddress,
        int intMaxMessages
    )
    {
        // Initialize EWS service.
        ExchangeService service = getService(emlUserAddress, apiKey);

        //sanitize request data
        intMaxMessages = Math.Max(1, Math.Min(100, intMaxMessages)); // 100>=NumMessages>=1

        Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);
        FindItemsResults<Item> messages = inbox.FindItems(new ItemView(intMaxMessages));

        return messages;
    }

    #endregion

    #region Calendar

    [WebMethod]
    public WhartonEWSResponse CreateCalendarItem
    (
        String apiKey,
        String emlUserAddress, 
        String lstReqAttendeeEmail,
        String lstOptAttendeeEmail,
        String dtCalItemStart, 
        String dtCalItemEnd, 
        String strCalItemSubject, 
        String strCalItemLocation, 
        String strCalItemBody,
        Boolean blnAllDayFlag,
        String strCalItemCategories
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            //parse input date/time values into native DateTime types
            DateTime rangeStart = DateTime.Parse(dtCalItemStart);
            DateTime rangeEnd = DateTime.Parse(dtCalItemEnd);

            //create the appointment object
            Appointment newAppointment = new Appointment(service);

            //fill in the details
            newAppointment.Subject = strCalItemSubject;
            newAppointment.Body = strCalItemBody;
            newAppointment.Body.BodyType = BodyType.Text;
            newAppointment.Start = rangeStart;
            newAppointment.End = rangeEnd;
            newAppointment.Location = strCalItemLocation;
            newAppointment.IsAllDayEvent = blnAllDayFlag;

            if (lstReqAttendeeEmail.Length > 0)
            {
                lstReqAttendeeEmail.Replace(';', ',');
                String[] reqAttendees = lstReqAttendeeEmail.Split(',');
                foreach (String a in reqAttendees)
                {
                    newAppointment.RequiredAttendees.Add(a.Trim());
                }
            }
            if (lstOptAttendeeEmail.Length > 0)
            {
                lstOptAttendeeEmail.Replace(';', ',');
                String[] optAttendees = lstOptAttendeeEmail.Split(',');
                foreach (String a in optAttendees)
                {
                    newAppointment.OptionalAttendees.Add(a.Trim());
                }
            }

            if (strCalItemCategories.Length > 0)
            {
                String[] cats = strCalItemCategories.Split(',');
                foreach (String cat in cats)
                {
                    newAppointment.Categories.Add(cat);
                }
            }

            //save it!
            newAppointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);

            rsp.SimpleData = true;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "CreateCalendarItem", success, miscData);

        return rsp;
    }

    //lookup method shared by two webmethods
    private FindItemsResults<Appointment> getCalItemsBase
    (
        String apiKey,
        String emlUserAddress,
        String dtRangeBegin,
        String dtRangeEnd
    )
    {
        // Initialize EWS service.
        ExchangeService service = getService(emlUserAddress, apiKey);

        //parse input date/time values into native DateTime types
        DateTime rangeStart = DateTime.Parse(dtRangeBegin);
        DateTime rangeEnd = DateTime.Parse(dtRangeEnd);

        ////Bind to the logged on user's calendar folder
        CalendarFolder myCalendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar);

        ////Query for items for the upcoming week - Expanding all recurrences 
        ////using CalendarView 
        return myCalendar.FindAppointments(new CalendarView(rangeStart, rangeEnd));

    }

    [WebMethod]
    public WhartonEWSResponse GetCalendarItems
    (
        String apiKey,
        String emlUserAddress,
        String dtRangeBegin,
        String dtRangeEnd
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            FindItemsResults<Appointment> myAppointments = getCalItemsBase(apiKey, emlUserAddress, dtRangeBegin, dtRangeEnd);

            // Create a DataTable. 
            DataTable table = new DataTable("Appointments");

            #region table column definitions

            DataColumn dcSubject = new DataColumn();
            dcSubject.DataType = System.Type.GetType("System.String");
            dcSubject.AllowDBNull = true;
            dcSubject.Caption = "Subject";
            dcSubject.ColumnName = "Subject";
            dcSubject.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcSubject);

            DataColumn dcLocation = new DataColumn();
            dcLocation.DataType = System.Type.GetType("System.String");
            dcLocation.AllowDBNull = true;
            dcLocation.Caption = "Location";
            dcLocation.ColumnName = "Location";
            dcLocation.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLocation);

            DataColumn dcLink = new DataColumn();
            dcLink.DataType = System.Type.GetType("System.String");
            dcLink.AllowDBNull = true;
            dcLink.Caption = "Link";
            dcLink.ColumnName = "Link";
            dcLink.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLink);

            DataColumn dcStart = new DataColumn();
            dcStart.DataType = System.Type.GetType("System.DateTime");
            dcStart.AllowDBNull = true;
            dcStart.Caption = "Start";
            dcStart.ColumnName = "Start";
            // Add the column to the table. 
            table.Columns.Add(dcStart);

            DataColumn dcEnd = new DataColumn();
            dcEnd.DataType = System.Type.GetType("System.DateTime");
            dcEnd.AllowDBNull = true;
            dcEnd.Caption = "End";
            dcEnd.ColumnName = "End";
            // Add the column to the table. 
            table.Columns.Add(dcEnd);

            DataColumn dcAllDay = new DataColumn();
            dcAllDay.DataType = System.Type.GetType("System.String");
            dcAllDay.AllowDBNull = true;
            dcAllDay.Caption = "AllDay";
            dcAllDay.ColumnName = "AllDay";
            dcAllDay.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcAllDay);

            DataColumn dcID = new DataColumn();
            dcID.DataType = System.Type.GetType("System.String");
            dcID.AllowDBNull = true;
            dcID.Caption = "Id";
            dcID.ColumnName = "Id";
            dcID.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcID);

            #endregion

            // Add rows and set values. 
            DataRow row;

            ////Write out the subject of each appointment I have this week
            foreach (Appointment appointment in myAppointments)
            {
                string openItemUrl = ItemUrlProtocol + service.Url.Host + ItemUrlFolder + appointment.WebClientReadFormQueryString;

                row = table.NewRow();
                row["Subject"] = cleanHighAsciiForXML(appointment.Subject);
                row["Location"] = cleanHighAsciiForXML(appointment.Location);
                row["Start"] = appointment.Start;
                row["End"] = appointment.End;
                row["Id"] = appointment.Id;

                try
                {
                    row["AllDay"] = appointment.IsAllDayEvent.ToString();
                }
                catch (Exception e)
                {
                    row["AllDay"] = false; //default to false
                }
                row["Link"] = openItemUrl;

                // Be sure to add the new row to the 
                // DataRowCollection. 
                table.Rows.Add(row);
            }

            rsp.TableData = table;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetCalendarItems", success, miscData);

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse GetCalendarItemsDetailed
    (
        String apiKey,
        String emlUserAddress,
        String dtRangeBegin,
        String dtRangeEnd
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            FindItemsResults<Appointment> myAppointments = getCalItemsBase(apiKey, emlUserAddress, dtRangeBegin, dtRangeEnd);

            // Create a DataTable. 
            DataTable table = new DataTable("Appointments");

            #region table column definitions

            DataColumn dcSubject = new DataColumn();
            dcSubject.DataType = System.Type.GetType("System.String");
            dcSubject.AllowDBNull = true;
            dcSubject.Caption = "Subject";
            dcSubject.ColumnName = "Subject";
            dcSubject.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcSubject);

            DataColumn dcLocation = new DataColumn();
            dcLocation.DataType = System.Type.GetType("System.String");
            dcLocation.AllowDBNull = true;
            dcLocation.Caption = "Location";
            dcLocation.ColumnName = "Location";
            dcLocation.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLocation);

            DataColumn dcLink = new DataColumn();
            dcLink.DataType = System.Type.GetType("System.String");
            dcLink.AllowDBNull = true;
            dcLink.Caption = "Link";
            dcLink.ColumnName = "Link";
            dcLink.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLink);

            DataColumn dcStart = new DataColumn();
            dcStart.DataType = System.Type.GetType("System.DateTime");
            dcStart.AllowDBNull = true;
            dcStart.Caption = "Start";
            dcStart.ColumnName = "Start";
            // Add the column to the table. 
            table.Columns.Add(dcStart);

            DataColumn dcEnd = new DataColumn();
            dcEnd.DataType = System.Type.GetType("System.DateTime");
            dcEnd.AllowDBNull = true;
            dcEnd.Caption = "End";
            dcEnd.ColumnName = "End";
            // Add the column to the table. 
            table.Columns.Add(dcEnd);

            DataColumn dcAllDay = new DataColumn();
            dcAllDay.DataType = System.Type.GetType("System.String");
            dcAllDay.AllowDBNull = true;
            dcAllDay.Caption = "AllDay";
            dcAllDay.ColumnName = "AllDay";
            dcAllDay.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcAllDay);

            DataColumn dcBody = new DataColumn();
            dcBody.DataType = System.Type.GetType("System.String");
            dcBody.AllowDBNull = true;
            dcBody.Caption = "Body";
            dcBody.ColumnName = "Body";
            dcBody.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcBody);

            DataColumn dcCategories = new DataColumn();
            dcCategories.DataType = System.Type.GetType("System.String");
            dcCategories.AllowDBNull = true;
            dcCategories.Caption = "Categories";
            dcCategories.ColumnName = "Categories";
            dcCategories.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcCategories);

            DataColumn dcReqAttendees = new DataColumn();
            dcReqAttendees.DataType = System.Type.GetType("System.String");
            dcReqAttendees.AllowDBNull = true;
            dcReqAttendees.Caption = "ReqAttendees";
            dcReqAttendees.ColumnName = "ReqAttendees";
            dcReqAttendees.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcReqAttendees);

            DataColumn dcOptAttendees = new DataColumn();
            dcOptAttendees.DataType = System.Type.GetType("System.String");
            dcOptAttendees.AllowDBNull = true;
            dcOptAttendees.Caption = "OptAttendees";
            dcOptAttendees.ColumnName = "OptAttendees";
            dcOptAttendees.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcOptAttendees);

            DataColumn dcId = new DataColumn();
            dcId.DataType = System.Type.GetType("System.String");
            dcId.AllowDBNull = true;
            dcId.Caption = "Id";
            dcId.ColumnName = "Id";
            dcId.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcId);

            #endregion

            // Add rows and set values. 
            DataRow row;

            //load properties
            List<Item> items = new List<Item>();
            if (myAppointments.Items.Count > 0)
            {
                foreach (Appointment a in myAppointments)
                {
                    items.Add(a);
                }
            }
            service.LoadPropertiesForItems(items, PropertySet.FirstClassProperties);

            ////Write out the subject of each appointment I have this week
            foreach (Appointment appointment in myAppointments)
            {
                string openItemUrl = ItemUrlProtocol + service.Url.Host + ItemUrlFolder + appointment.WebClientReadFormQueryString;

                row = table.NewRow();
                row["Subject"] = cleanHighAsciiForXML(appointment.Subject);
                row["Location"] = cleanHighAsciiForXML(appointment.Location);
                row["Start"] = appointment.Start;
                row["End"] = appointment.End;
                row["Body"] = appointment.Body.Text;
                row["Categories"] = appointment.Categories;
                row["Id"] = appointment.Id;
                
                //required attendees
                AttendeeCollection req = appointment.RequiredAttendees;
                StringList tmp = new StringList();
                foreach (Attendee att in req)
                {
                    tmp.Add(att.Address);
                }
                row["ReqAttendees"] = tmp.ToString();

                //optional attendees
                AttendeeCollection opt = appointment.OptionalAttendees;
                tmp = new StringList();
                foreach (Attendee att in opt)
                {
                    tmp.Add(att.Address);
                }
                row["OptAttendees"] = tmp.ToString();

                try
                {
                    row["AllDay"] = appointment.IsAllDayEvent.ToString();
                }
                catch (Exception e)
                {
                    row["AllDay"] = false; //default to false
                }
                row["Link"] = openItemUrl;

                // Be sure to add the new row to the 
                // DataRowCollection. 
                table.Rows.Add(row);
            }

            rsp.TableData = table;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetCalendarItems", success, miscData);

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse GetCalendarItemById
    (
        String apiKey,
        String strUniqueId,
        String emlUserAddress
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            Appointment appointment = GetAppointment(service, strUniqueId);

            #region table definitions

            // Create a DataTable. 
            DataTable table = new DataTable("Appointments");

            DataColumn dcSubject = new DataColumn();
            dcSubject.DataType = System.Type.GetType("System.String");
            dcSubject.AllowDBNull = true;
            dcSubject.Caption = "Subject";
            dcSubject.ColumnName = "Subject";
            dcSubject.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcSubject);

            DataColumn dcLocation = new DataColumn();
            dcLocation.DataType = System.Type.GetType("System.String");
            dcLocation.AllowDBNull = true;
            dcLocation.Caption = "Location";
            dcLocation.ColumnName = "Location";
            dcLocation.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLocation);

            DataColumn dcLink = new DataColumn();
            dcLink.DataType = System.Type.GetType("System.String");
            dcLink.AllowDBNull = true;
            dcLink.Caption = "Link";
            dcLink.ColumnName = "Link";
            dcLink.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLink);

            DataColumn dcStart = new DataColumn();
            dcStart.DataType = System.Type.GetType("System.DateTime");
            dcStart.AllowDBNull = true;
            dcStart.Caption = "Start";
            dcStart.ColumnName = "Start";
            // Add the column to the table. 
            table.Columns.Add(dcStart);

            DataColumn dcEnd = new DataColumn();
            dcEnd.DataType = System.Type.GetType("System.DateTime");
            dcEnd.AllowDBNull = true;
            dcEnd.Caption = "End";
            dcEnd.ColumnName = "End";
            // Add the column to the table. 
            table.Columns.Add(dcEnd);

            DataColumn dcAllDay = new DataColumn();
            dcAllDay.DataType = System.Type.GetType("System.String");
            dcAllDay.AllowDBNull = true;
            dcAllDay.Caption = "AllDay";
            dcAllDay.ColumnName = "AllDay";
            dcAllDay.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcAllDay);

            DataColumn dcBody = new DataColumn();
            dcBody.DataType = System.Type.GetType("System.String");
            dcBody.AllowDBNull = true;
            dcBody.Caption = "Body";
            dcBody.ColumnName = "Body";
            dcBody.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcBody);

            DataColumn dcCategories = new DataColumn();
            dcCategories.DataType = System.Type.GetType("System.String");
            dcCategories.AllowDBNull = true;
            dcCategories.Caption = "Categories";
            dcCategories.ColumnName = "Categories";
            dcCategories.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcCategories);

            DataColumn dcReqAttendees = new DataColumn();
            dcReqAttendees.DataType = System.Type.GetType("System.String");
            dcReqAttendees.AllowDBNull = true;
            dcReqAttendees.Caption = "ReqAttendees";
            dcReqAttendees.ColumnName = "ReqAttendees";
            dcReqAttendees.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcReqAttendees);

            DataColumn dcOptAttendees = new DataColumn();
            dcOptAttendees.DataType = System.Type.GetType("System.String");
            dcOptAttendees.AllowDBNull = true;
            dcOptAttendees.Caption = "OptAttendees";
            dcOptAttendees.ColumnName = "OptAttendees";
            dcOptAttendees.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcOptAttendees);

            #endregion

            #region load details

            //load properties
            List<Item> items = new List<Item>();
            items.Add(appointment);
            service.LoadPropertiesForItems(items, PropertySet.FirstClassProperties);

            #endregion

            string openItemUrl = ItemUrlProtocol + service.Url.Host + ItemUrlFolder + appointment.WebClientReadFormQueryString;

            // Add rows and set values. 
            DataRow row;

            row = table.NewRow();
            row["Subject"] = cleanHighAsciiForXML(appointment.Subject);
            row["Location"] = cleanHighAsciiForXML(appointment.Location);
            row["Start"] = appointment.Start;
            row["End"] = appointment.End;
            row["Body"] = appointment.Body.Text;
            row["Categories"] = appointment.Categories;

            //required attendees
            AttendeeCollection req = appointment.RequiredAttendees;
            StringList tmp = new StringList();
            foreach (Attendee att in req)
            {
                tmp.Add(att.Address);
            }
            row["ReqAttendees"] = tmp.ToString();

            //optional attendees
            AttendeeCollection opt = appointment.OptionalAttendees;
            tmp = new StringList();
            foreach (Attendee att in opt)
            {
                tmp.Add(att.Address);
            }
            row["OptAttendees"] = tmp.ToString();

            try
            {
                row["AllDay"] = appointment.IsAllDayEvent.ToString();
            }
            catch (Exception e)
            {
                row["AllDay"] = false; //default to false
            }
            row["Link"] = openItemUrl;

            // Be sure to add the new row to the 
            // DataRowCollection. 
            table.Rows.Add(row);

            rsp.TableData = table;

        }
        catch (Exception e)
        {
            if (e.Message == "The specified object was not found in the store.")
            {
                rsp.Msg = e.Message;
                rsp.StatusCode = 404;
            }
            else
            {
                if (!(int.TryParse(e.Message, out rsp.StatusCode)))
                {
                    //if the parsing fails, then set a default value of 500
                    rsp.StatusCode = 500;
                }
                rsp.Msg = e.Message;
            }

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\", \"UniqueId\":\"" + strUniqueId + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetCalendarItemById", success, miscData);

        rsp.SimpleData = success;

        return rsp;
    }

    private Appointment GetAppointment
    (
        ExchangeService service,
        String uniqueId
    )
    {
        Appointment a = Appointment.Bind(service, uniqueId);
        return a;
    }

    [WebMethod]
    public WhartonEWSResponse UpdateCalendarItem
    (
        String apiKey,
        String strUniqueId,
        String emlUserAddress,
        String lstReqAttendeeEmail,
        String lstOptAttendeeEmail,
        String dtCalItemStart,
        String dtCalItemEnd,
        String strCalItemSubject,
        String strCalItemLocation,
        String strCalItemBody,
        Boolean blnAllDayFlag,
        String strCalItemCategories
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            Appointment a = GetAppointment(service, strUniqueId);

            //parse input date/time values into native DateTime types
            DateTime rangeStart = DateTime.Parse(dtCalItemStart);
            DateTime rangeEnd = DateTime.Parse(dtCalItemEnd);

            //create the appointment object
            Appointment newAppointment = new Appointment(service);

            //fill in the details
            a.Subject = strCalItemSubject;
            a.Body = strCalItemBody;
            a.Body.BodyType = BodyType.Text;
            a.Start = rangeStart;
            a.End = rangeEnd;
            a.Location = strCalItemLocation;
            a.IsAllDayEvent = blnAllDayFlag;

            if (lstReqAttendeeEmail.Length > 0)
            {
                lstReqAttendeeEmail.Replace(';', ',');
                String[] reqAttendees = lstReqAttendeeEmail.Split(',');
                foreach (String att in reqAttendees)
                {
                    a.RequiredAttendees.Add(att.Trim());
                }
            }
            if (lstOptAttendeeEmail.Length > 0)
            {
                lstOptAttendeeEmail.Replace(';', ',');
                String[] optAttendees = lstOptAttendeeEmail.Split(',');
                foreach (String att in optAttendees)
                {
                    a.OptionalAttendees.Add(att.Trim());
                }
            }

            if (strCalItemCategories.Length > 0)
            {
                String[] cats = strCalItemCategories.Split(',');
                foreach (String cat in cats)
                {
                    a.Categories.Add(cat);
                }
            }

            //save it!
            a.Update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendOnlyToAll);

            rsp.SimpleData = true;

        }
        catch (Exception e)
        {
            if (e.Message == "The specified object was not found in the store.")
            {
                rsp.Msg = e.Message;
                rsp.StatusCode = 404;
            }
            else
            {
                if (!(int.TryParse(e.Message, out rsp.StatusCode)))
                {
                    //if the parsing fails, then set a default value of 500
                    rsp.StatusCode = 500;
                }
                rsp.Msg = e.Message;
            }

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\", \"UniqueId\":\"" + strUniqueId + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetCalendarItemById", success, miscData);

        rsp.SimpleData = success;

        return rsp;
    }

    #endregion

    #region Contacts

    [WebMethod]
    public WhartonEWSResponse CreateContact
    (
        String apiKey,
        String emlUserAddress, 
        String strCtctGivenName,
        String strCtctSurname,
        String strCtctCompanyName,
        String strCtctHomePhone,
        String strCtctWorkPhone,
        String strCtctMobilePhone,
        String emlCtctEmail1,
        String emlCtctEmail2,
        String emlCtctEmail3,
        String strCtctHomeAddrStreet,
        String strCtctHomeAddrCity,
        String strCtctHomeAddrStateAbbr,
        String strCtctHomeAddrPostalCode,
        String strCtctHomeAddrCountry,
        String strCtctWorkAddrStreet,
        String strCtctWorkAddrCity,
        String strCtctWorkAddrStateAbbr,
        String strCtctWorkAddrPostalCode,
        String strCtctWorkAddrCountry
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            Contact c = new Contact(service);

            c.GivenName = strCtctGivenName;
            c.Surname = strCtctSurname;
            c.FileAsMapping = FileAsMapping.SurnameCommaGivenName;

            //add company name if supplied
            if (strCtctCompanyName.Length > 0)
                c.CompanyName = strCtctCompanyName;

            //add suppied phone numbers
            if (strCtctHomePhone.Length > 0)
                c.PhoneNumbers[PhoneNumberKey.HomePhone] = strCtctHomePhone;
            if (strCtctWorkPhone.Length > 0)
                c.PhoneNumbers[PhoneNumberKey.BusinessPhone] = strCtctWorkPhone;
            if (strCtctMobilePhone.Length > 0)
                c.PhoneNumbers[PhoneNumberKey.MobilePhone] = strCtctMobilePhone;

            //add supplied email addresses
            if (emlCtctEmail1.Length > 0)
                c.EmailAddresses[EmailAddressKey.EmailAddress1] = new EmailAddress(emlCtctEmail1);
            if (emlCtctEmail2.Length > 0)
                c.EmailAddresses[EmailAddressKey.EmailAddress2] = new EmailAddress(emlCtctEmail2);
            if (emlCtctEmail3.Length > 0)
                c.EmailAddresses[EmailAddressKey.EmailAddress3] = new EmailAddress(emlCtctEmail3);

            //Home address
            int homeAddrLength = 0;
            homeAddrLength += strCtctHomeAddrStreet.Length;
            homeAddrLength += strCtctHomeAddrCity.Length;
            homeAddrLength += strCtctHomeAddrStateAbbr.Length;
            homeAddrLength += strCtctHomeAddrPostalCode.Length;
            homeAddrLength += strCtctHomeAddrCountry.Length;
            if (homeAddrLength > 0)
            {
                PhysicalAddressEntry paEntry1 = new PhysicalAddressEntry();

                //add whichever fields are supplied
                if (strCtctHomeAddrStreet.Length > 0)
                    paEntry1.Street = strCtctHomeAddrStreet;
                if (strCtctHomeAddrCity.Length > 0)
                    paEntry1.City = strCtctHomeAddrCity;
                if (strCtctHomeAddrStateAbbr.Length > 0)
                    paEntry1.State = strCtctHomeAddrStateAbbr;
                if (strCtctHomeAddrPostalCode.Length > 0)
                    paEntry1.PostalCode = strCtctHomeAddrPostalCode;
                if (strCtctHomeAddrCountry.Length > 0)
                    paEntry1.CountryOrRegion = strCtctHomeAddrCountry;

                c.PhysicalAddresses[PhysicalAddressKey.Home] = paEntry1;
            }

            //Work address
            int workAddrLength = 0;
            workAddrLength += strCtctWorkAddrStreet.Length;
            workAddrLength += strCtctWorkAddrCity.Length;
            workAddrLength += strCtctWorkAddrStateAbbr.Length;
            workAddrLength += strCtctWorkAddrPostalCode.Length;
            workAddrLength += strCtctWorkAddrCountry.Length;
            if (workAddrLength > 0)
            {
                PhysicalAddressEntry paEntry2 = new PhysicalAddressEntry();

                //add whichever fields are supplied
                if (strCtctWorkAddrStreet.Length > 0)
                    paEntry2.Street = strCtctWorkAddrStreet;
                if (strCtctWorkAddrCity.Length > 0)
                    paEntry2.City = strCtctWorkAddrCity;
                if (strCtctWorkAddrStateAbbr.Length > 0)
                    paEntry2.State = strCtctWorkAddrStateAbbr;
                if (strCtctWorkAddrPostalCode.Length > 0)
                    paEntry2.PostalCode = strCtctWorkAddrPostalCode;
                if (strCtctWorkAddrCountry.Length > 0)
                    paEntry2.CountryOrRegion = strCtctWorkAddrCountry;

                c.PhysicalAddresses[PhysicalAddressKey.Business] = paEntry2;
            }

            //save the contact
            c.Save(WellKnownFolderName.Contacts);

            rsp.SimpleData = c.Id.UniqueId;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "CreateContact", success, miscData);

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse UpdateContact(
        String apiKey,
        String strUniqueId,
        String emlUserAddress,
        String strCtctGivenName,
        String strCtctSurname,
        String strCtctCompanyName,
        String strCtctHomePhone,
        String strCtctWorkPhone,
        String strCtctMobilePhone,
        String emlCtctEmail1,
        String emlCtctEmail2,
        String emlCtctEmail3,
        String strCtctHomeAddrStreet,
        String strCtctHomeAddrCity,
        String strCtctHomeAddrStateAbbr,
        String strCtctHomeAddrPostalCode,
        String strCtctHomeAddrCountry,
        String strCtctWorkAddrStreet,
        String strCtctWorkAddrCity,
        String strCtctWorkAddrStateAbbr,
        String strCtctWorkAddrPostalCode,
        String strCtctWorkAddrCountry
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        ExchangeService service = getService(emlUserAddress, apiKey);

        try
        {
            Contact c = Contact.Bind(service, new ItemId(strUniqueId));

            c.GivenName = strCtctGivenName;
            c.Surname = strCtctSurname;
            c.FileAsMapping = FileAsMapping.SurnameCommaGivenName;

            //add company name if supplied
            if (strCtctCompanyName.Length > 0)
                c.CompanyName = strCtctCompanyName;
            else
                c.CompanyName = null;

            //add suppied phone numbers
            if (strCtctHomePhone.Length > 0)
                c.PhoneNumbers[PhoneNumberKey.HomePhone] = strCtctHomePhone;
            else
                c.PhoneNumbers[PhoneNumberKey.HomePhone] = null;

            if (strCtctWorkPhone.Length > 0)
                c.PhoneNumbers[PhoneNumberKey.BusinessPhone] = strCtctWorkPhone;
            else
                c.PhoneNumbers[PhoneNumberKey.BusinessPhone] = null;

            if (strCtctMobilePhone.Length > 0)
                c.PhoneNumbers[PhoneNumberKey.MobilePhone] = strCtctMobilePhone;
            else
                c.PhoneNumbers[PhoneNumberKey.MobilePhone] = null;

            //add supplied email addresses
            if (emlCtctEmail1.Length > 0)
                c.EmailAddresses[EmailAddressKey.EmailAddress1] = new EmailAddress(emlCtctEmail1);
            else
                c.EmailAddresses[EmailAddressKey.EmailAddress1] = null;

            if (emlCtctEmail2.Length > 0)
                c.EmailAddresses[EmailAddressKey.EmailAddress2] = new EmailAddress(emlCtctEmail2);
            else
                c.EmailAddresses[EmailAddressKey.EmailAddress2] = null;

            if (emlCtctEmail3.Length > 0)
                c.EmailAddresses[EmailAddressKey.EmailAddress3] = new EmailAddress(emlCtctEmail3);
            else
                c.EmailAddresses[EmailAddressKey.EmailAddress3] = null;

            //Home address
            int homeAddrLength = 0;
            homeAddrLength += strCtctHomeAddrStreet.Length;
            homeAddrLength += strCtctHomeAddrCity.Length;
            homeAddrLength += strCtctHomeAddrStateAbbr.Length;
            homeAddrLength += strCtctHomeAddrPostalCode.Length;
            homeAddrLength += strCtctHomeAddrCountry.Length;
            if (homeAddrLength > 0)
            {
                PhysicalAddressEntry paEntry1 = new PhysicalAddressEntry();

                //add whichever fields are supplied
                if (strCtctHomeAddrStreet.Length > 0)
                    paEntry1.Street = strCtctHomeAddrStreet;
                else
                    paEntry1.Street = null;

                if (strCtctHomeAddrCity.Length > 0)
                    paEntry1.City = strCtctHomeAddrCity;
                else
                    paEntry1.City = null;

                if (strCtctHomeAddrStateAbbr.Length > 0)
                    paEntry1.State = strCtctHomeAddrStateAbbr;
                else
                    paEntry1.State = null;

                if (strCtctHomeAddrPostalCode.Length > 0)
                    paEntry1.PostalCode = strCtctHomeAddrPostalCode;
                else
                    paEntry1.PostalCode = null;

                if (strCtctHomeAddrCountry.Length > 0)
                    paEntry1.CountryOrRegion = strCtctHomeAddrCountry;
                else
                    paEntry1.CountryOrRegion = null;

                c.PhysicalAddresses[PhysicalAddressKey.Home] = paEntry1;
            }
            else
            {
                c.PhysicalAddresses[PhysicalAddressKey.Home] = null;
            }

            //Work address
            int workAddrLength = 0;
            workAddrLength += strCtctWorkAddrStreet.Length;
            workAddrLength += strCtctWorkAddrCity.Length;
            workAddrLength += strCtctWorkAddrStateAbbr.Length;
            workAddrLength += strCtctWorkAddrPostalCode.Length;
            workAddrLength += strCtctWorkAddrCountry.Length;
            if (workAddrLength > 0)
            {
                PhysicalAddressEntry paEntry2 = new PhysicalAddressEntry();

                //add whichever fields are supplied
                if (strCtctWorkAddrStreet.Length > 0)
                    paEntry2.Street = strCtctWorkAddrStreet;
                else
                    paEntry2.Street = null;

                if (strCtctWorkAddrCity.Length > 0)
                    paEntry2.City = strCtctWorkAddrCity;
                else
                    paEntry2.City = null;

                if (strCtctWorkAddrStateAbbr.Length > 0)
                    paEntry2.State = strCtctWorkAddrStateAbbr;
                else
                    paEntry2.State = null;

                if (strCtctWorkAddrPostalCode.Length > 0)
                    paEntry2.PostalCode = strCtctWorkAddrPostalCode;
                else
                    paEntry2.PostalCode = null;

                if (strCtctWorkAddrCountry.Length > 0)
                    paEntry2.CountryOrRegion = strCtctWorkAddrCountry;
                else
                    paEntry2.CountryOrRegion = null;

                c.PhysicalAddresses[PhysicalAddressKey.Business] = paEntry2;
            }
            else
            {
                c.PhysicalAddresses[PhysicalAddressKey.Business] = null;
            }

            //save the contact
            c.Update(ConflictResolutionMode.AlwaysOverwrite);

            rsp.SimpleData = true;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "CreateContact", success, miscData);

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse DeleteContact
    (
        String apiKey,
        String strUniqueId,
        String emlUserAddress
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            Contact contact = Contact.Bind(service, new ItemId(strUniqueId));

            contact.Delete(DeleteMode.MoveToDeletedItems);
        }
        catch (Exception e)
        {
            if (e.Message == "The specified object was not found in the store.")
            {
                rsp.Msg = e.Message;
                rsp.StatusCode = 404;
            }
            else
            {
                if (!(int.TryParse(e.Message, out rsp.StatusCode)))
                {
                    //if the parsing fails, then set a default value of 500
                    rsp.StatusCode = 500;
                }
                rsp.Msg = e.Message;
            }

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\", \"UniqueId\":\"" + strUniqueId + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "DeleteContact", success, miscData);

        rsp.SimpleData = success;

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse GetContacts
    (
        String apiKey,
        String emlUserAddress
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            //config
            int maxContacts = 10000;

            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            //Bind to the logged on user's calendar folder
            ContactsFolder myContacts = ContactsFolder.Bind(service, WellKnownFolderName.Contacts);

            //Query for items
            FindItemsResults<Item> contactList = myContacts.FindItems(new ItemView(maxContacts));

            // Create a DataTable. 
            DataTable table = new DataTable("Contacts");

            #region table column definitions

            DataColumn dcFileAs = new DataColumn();
            dcFileAs.DataType = System.Type.GetType("System.String");
            dcFileAs.AllowDBNull = true;
            dcFileAs.Caption = "FileAs";
            dcFileAs.ColumnName = "FileAs";
            dcFileAs.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcFileAs);

            DataColumn dcLink = new DataColumn();
            dcLink.DataType = System.Type.GetType("System.String");
            dcLink.AllowDBNull = true;
            dcLink.Caption = "Link";
            dcLink.ColumnName = "Link";
            dcLink.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLink);

            DataColumn dcGivenName = new DataColumn();
            dcGivenName.DataType = System.Type.GetType("System.String");
            dcGivenName.AllowDBNull = true;
            dcGivenName.Caption = "GivenName";
            dcGivenName.ColumnName = "GivenName";
            dcGivenName.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcGivenName);

            DataColumn dcSurname = new DataColumn();
            dcSurname.DataType = System.Type.GetType("System.String");
            dcSurname.AllowDBNull = true;
            dcSurname.Caption = "Surname";
            dcSurname.ColumnName = "Surname";
            // Add the column to the table. 
            table.Columns.Add(dcSurname);

            DataColumn dcCompanyName = new DataColumn();
            dcCompanyName.DataType = System.Type.GetType("System.String");
            dcCompanyName.AllowDBNull = true;
            dcCompanyName.Caption = "CompanyName";
            dcCompanyName.ColumnName = "CompanyName";
            // Add the column to the table. 
            table.Columns.Add(dcCompanyName);

            DataColumn dcEmails = new DataColumn();
            dcEmails.DataType = System.Type.GetType("System.String[]");
            dcEmails.AllowDBNull = true;
            dcEmails.Caption = "Emails";
            dcEmails.ColumnName = "Emails";
            // Add the column to the table. 
            table.Columns.Add(dcEmails);

            DataColumn dcMobilePhone = new DataColumn();
            dcMobilePhone.DataType = System.Type.GetType("System.String");
            dcMobilePhone.AllowDBNull = true;
            dcMobilePhone.Caption = "MobilePhone";
            dcMobilePhone.ColumnName = "MobilePhone";
            // Add the column to the table. 
            table.Columns.Add(dcMobilePhone);

            DataColumn dcBusPhone = new DataColumn();
            dcBusPhone.DataType = System.Type.GetType("System.String");
            dcBusPhone.AllowDBNull = true;
            dcBusPhone.Caption = "BusPhone";
            dcBusPhone.ColumnName = "BusPhone";
            // Add the column to the table. 
            table.Columns.Add(dcBusPhone);

            DataColumn dcHomePhone = new DataColumn();
            dcHomePhone.DataType = System.Type.GetType("System.String");
            dcHomePhone.AllowDBNull = true;
            dcHomePhone.Caption = "HomePhone";
            dcHomePhone.ColumnName = "HomePhone";
            // Add the column to the table. 
            table.Columns.Add(dcHomePhone);

            DataColumn dcHomeStreet = new DataColumn();
            dcHomeStreet.DataType = System.Type.GetType("System.String");
            dcHomeStreet.AllowDBNull = true;
            dcHomeStreet.Caption = "HomeStreet";
            dcHomeStreet.ColumnName = "HomeStreet";
            // Add the column to the table. 
            table.Columns.Add(dcHomeStreet);

            DataColumn dcHomeCity = new DataColumn();
            dcHomeCity.DataType = System.Type.GetType("System.String");
            dcHomeCity.AllowDBNull = true;
            dcHomeCity.Caption = "HomeCity";
            dcHomeCity.ColumnName = "HomeCity";
            // Add the column to the table. 
            table.Columns.Add(dcHomeCity);

            DataColumn dcHomeState = new DataColumn();
            dcHomeState.DataType = System.Type.GetType("System.String");
            dcHomeState.AllowDBNull = true;
            dcHomeState.Caption = "HomeState";
            dcHomeState.ColumnName = "HomeState";
            // Add the column to the table. 
            table.Columns.Add(dcHomeState);

            DataColumn dcHomeZip = new DataColumn();
            dcHomeZip.DataType = System.Type.GetType("System.String");
            dcHomeZip.AllowDBNull = true;
            dcHomeZip.Caption = "HomeZip";
            dcHomeZip.ColumnName = "HomeZip";
            // Add the column to the table. 
            table.Columns.Add(dcHomeZip);

            DataColumn dcHomeCountry = new DataColumn();
            dcHomeCountry.DataType = System.Type.GetType("System.String");
            dcHomeCountry.AllowDBNull = true;
            dcHomeCountry.Caption = "HomeCountry";
            dcHomeCountry.ColumnName = "HomeCountry";
            // Add the column to the table. 
            table.Columns.Add(dcHomeCountry);

            DataColumn dcWorkStreet = new DataColumn();
            dcWorkStreet.DataType = System.Type.GetType("System.String");
            dcWorkStreet.AllowDBNull = true;
            dcWorkStreet.Caption = "WorkStreet";
            dcWorkStreet.ColumnName = "WorkStreet";
            // Add the column to the table. 
            table.Columns.Add(dcWorkStreet);

            DataColumn dcWorkCity = new DataColumn();
            dcWorkCity.DataType = System.Type.GetType("System.String");
            dcWorkCity.AllowDBNull = true;
            dcWorkCity.Caption = "WorkCity";
            dcWorkCity.ColumnName = "WorkCity";
            // Add the column to the table. 
            table.Columns.Add(dcWorkCity);

            DataColumn dcWorkState = new DataColumn();
            dcWorkState.DataType = System.Type.GetType("System.String");
            dcWorkState.AllowDBNull = true;
            dcWorkState.Caption = "WorkState";
            dcWorkState.ColumnName = "WorkState";
            // Add the column to the table. 
            table.Columns.Add(dcWorkState);

            DataColumn dcWorkZip = new DataColumn();
            dcWorkZip.DataType = System.Type.GetType("System.String");
            dcWorkZip.AllowDBNull = true;
            dcWorkZip.Caption = "WorkZip";
            dcWorkZip.ColumnName = "WorkZip";
            // Add the column to the table. 
            table.Columns.Add(dcWorkZip);

            DataColumn dcWorkCountry = new DataColumn();
            dcWorkCountry.DataType = System.Type.GetType("System.String");
            dcWorkCountry.AllowDBNull = true;
            dcWorkCountry.Caption = "WorkCountry";
            dcWorkCountry.ColumnName = "WorkCountry";
            // Add the column to the table. 
            table.Columns.Add(dcWorkCountry);

            DataColumn dcUniqueId = new DataColumn();
            dcUniqueId.DataType = System.Type.GetType("System.String");
            dcUniqueId.AllowDBNull = true;
            dcUniqueId.Caption = "UniqueId";
            dcUniqueId.ColumnName = "UniqueId";
            // Add the column to the table. 
            table.Columns.Add(dcUniqueId);

            #endregion

            //build table of contact data
            DataRow row;
            foreach (Object o in contactList)
            {
                if (o.GetType().Name.ToString() != "Contact")
                {
                    continue; //skip this record, it's a distribution list or something like that
                }

                Contact c = (Contact)o;

                string openItemUrl = ItemUrlProtocol + service.Url.Host + ItemUrlFolder + c.WebClientReadFormQueryString;

                row = table.NewRow();
                row["FileAs"] = cleanHighAsciiForXML(c.FileAs);
                row["GivenName"] = cleanHighAsciiForXML(c.GivenName);
                row["Surname"] = cleanHighAsciiForXML(c.Surname);
                row["CompanyName"] = cleanHighAsciiForXML(c.CompanyName);
                row["Link"] = openItemUrl;
                row["UniqueId"] = c.Id.UniqueId;

                //home address
                PhysicalAddressEntry addressEntry;
                PhysicalAddressKey addressKey = PhysicalAddressKey.Home;
                if (c.PhysicalAddresses.TryGetValue(addressKey, out addressEntry))
                {
                    if (addressEntry.Street != null)
                    {
                        row["HomeStreet"] = cleanHighAsciiForXML(addressEntry.Street.ToString());
                    }
                    if (addressEntry.City != null)
                    {
                        row["HomeCity"] = cleanHighAsciiForXML(addressEntry.City.ToString());
                    }
                    if (addressEntry.State != null)
                    {
                        row["HomeState"] = cleanHighAsciiForXML(addressEntry.State.ToString());
                    }
                    if (addressEntry.PostalCode != null)
                    {
                        row["HomeZip"] = cleanHighAsciiForXML(addressEntry.PostalCode.ToString());
                    }
                    if (addressEntry.CountryOrRegion != null)
                    {
                        row["HomeCountry"] = cleanHighAsciiForXML(addressEntry.CountryOrRegion.ToString());
                    }
                }

                //work address
                addressEntry = null;
                addressKey = PhysicalAddressKey.Business;
                if (c.PhysicalAddresses.TryGetValue(addressKey, out addressEntry))
                {
                    if (addressEntry.Street != null)
                    {
                        row["WorkStreet"] = cleanHighAsciiForXML(addressEntry.Street.ToString());
                    }
                    if (addressEntry.City != null)
                    {
                        row["WorkCity"] = cleanHighAsciiForXML(addressEntry.City.ToString());
                    }
                    if (addressEntry.State != null)
                    {
                        row["WorkState"] = cleanHighAsciiForXML(addressEntry.State.ToString());
                    }
                    if (addressEntry.PostalCode != null)
                    {
                        row["WorkZip"] = cleanHighAsciiForXML(addressEntry.PostalCode.ToString());
                    }
                    if (addressEntry.CountryOrRegion != null)
                    {
                        row["WorkCountry"] = cleanHighAsciiForXML(addressEntry.CountryOrRegion.ToString());
                    }
                }

                //phone numbers
                String phoneEntry = null;
                PhoneNumberKey phoneKey = PhoneNumberKey.HomePhone;
                if (c.PhoneNumbers.TryGetValue(phoneKey, out phoneEntry))
                {
                    row["HomePhone"] = cleanHighAsciiForXML(phoneEntry);
                }
                phoneEntry = null;
                phoneKey = PhoneNumberKey.BusinessPhone;
                if (c.PhoneNumbers.TryGetValue(phoneKey, out phoneEntry))
                {
                    row["BusPhone"] = cleanHighAsciiForXML(phoneEntry);
                }
                phoneEntry = null;
                phoneKey = PhoneNumberKey.MobilePhone;
                if (c.PhoneNumbers.TryGetValue(phoneKey, out phoneEntry))
                {
                    row["MobilePhone"] = cleanHighAsciiForXML(phoneEntry);
                }

                //emails
                String[] Emails = new String[3];
                EmailAddress emailEntry = null;
                EmailAddressKey emailKey = EmailAddressKey.EmailAddress1;
                if (c.EmailAddresses.TryGetValue(emailKey, out emailEntry))
                {
                    Emails[0] = cleanHighAsciiForXML(emailEntry.ToString());
                }
                emailEntry = null;
                emailKey = EmailAddressKey.EmailAddress2;
                if (c.EmailAddresses.TryGetValue(emailKey, out emailEntry))
                {
                    Emails[1] = cleanHighAsciiForXML(emailEntry.ToString());
                }
                emailEntry = null;
                emailKey = EmailAddressKey.EmailAddress3;
                if (c.EmailAddresses.TryGetValue(emailKey, out emailEntry))
                {
                    Emails[2] = cleanHighAsciiForXML(emailEntry.ToString());
                }
                row["Emails"] = Emails;

                table.Rows.Add(row);
            }

            rsp.TableData = table;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetContacts", success, miscData);

        return rsp;
    }

    #endregion

    #region Tasks

    [WebMethod]
    public WhartonEWSResponse CreateTask
    (
        String apiKey,
        String emlUserAddress,
        String strTaskSubject,
        String strTaskStartDate,
        String strTaskDueDate,
        String strTaskImportance,
        String strTaskStatus,
        String strTaskBody
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            //parse task dates
            DateTime realStartDate = DateTime.Parse(strTaskStartDate);
            DateTime realDueDate = DateTime.Parse(strTaskDueDate);

            //parse task importance
            Importance realTskImportance;
            if (strTaskImportance.ToLower() == "high")
                realTskImportance = Importance.High;
            else if (strTaskImportance.ToLower() == "low")
                realTskImportance = Importance.Low;
            else
                realTskImportance = Importance.Normal;

            //parse task status
            TaskStatus realTskStatus;
            if (strTaskStatus.ToLower() == "inprogress")
                realTskStatus = TaskStatus.InProgress;
            else if (strTaskStatus.ToLower() == "completed")
                realTskStatus = TaskStatus.Completed;
            else if (strTaskStatus.ToLower() == "deferred")
                realTskStatus = TaskStatus.Deferred;
            else if (strTaskStatus.ToLower() == "waitingonothers")
                realTskStatus = TaskStatus.WaitingOnOthers;
            else
                realTskStatus = TaskStatus.NotStarted;

            //compose task object
            Task t = new Task(service);
            t.Subject = strTaskSubject;
            t.StartDate = realStartDate;
            t.DueDate = realDueDate;
            t.Importance = realTskImportance;
            t.Status = realTskStatus;
            t.Body = strTaskBody;

            //save to exchange
            t.Save(WellKnownFolderName.Tasks);

            rsp.SimpleData = true;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "CreateTask", success, miscData);

        return rsp;
    }

    [WebMethod]
    public WhartonEWSResponse GetTasks
    (
        String apiKey,
        String emlUserAddress
    )
    {
        WhartonEWSResponse rsp = new WhartonEWSResponse();

        Boolean success = true;
        String miscData = "";

        try
        {
            //config
            int maxTasks = 10000;

            // Initialize EWS service.
            ExchangeService service = getService(emlUserAddress, apiKey);

            //Bind to the logged on user's tasks folder
            TasksFolder myTasks = TasksFolder.Bind(service, WellKnownFolderName.Tasks);

            //Query for items
            FindItemsResults<Item> taskList = myTasks.FindItems(new ItemView(maxTasks));

            DataTable table = new DataTable("Tasks");

            #region table column definitions

            DataColumn dcSubject = new DataColumn();
            dcSubject.DataType = System.Type.GetType("System.String");
            dcSubject.AllowDBNull = true;
            dcSubject.Caption = "Subject";
            dcSubject.ColumnName = "Subject";
            dcSubject.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcSubject);

            DataColumn dcCreated = new DataColumn();
            dcCreated.DataType = System.Type.GetType("System.String");
            dcCreated.AllowDBNull = true;
            dcCreated.Caption = "Created";
            dcCreated.ColumnName = "Created";
            dcCreated.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcCreated);

            DataColumn dcStartDate = new DataColumn();
            dcStartDate.DataType = System.Type.GetType("System.String");
            dcStartDate.AllowDBNull = true;
            dcStartDate.Caption = "StartDate";
            dcStartDate.ColumnName = "StartDate";
            dcStartDate.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcStartDate);

            DataColumn dcDueDate = new DataColumn();
            dcDueDate.DataType = System.Type.GetType("System.String");
            dcDueDate.AllowDBNull = true;
            dcDueDate.Caption = "DueDate";
            dcDueDate.ColumnName = "DueDate";
            dcDueDate.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcDueDate);

            DataColumn dcImportance = new DataColumn();
            dcImportance.DataType = System.Type.GetType("System.String");
            dcImportance.AllowDBNull = true;
            dcImportance.Caption = "Importance";
            dcImportance.ColumnName = "Importance";
            dcImportance.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcImportance);

            DataColumn dcStatus = new DataColumn();
            dcStatus.DataType = System.Type.GetType("System.String");
            dcStatus.AllowDBNull = true;
            dcStatus.Caption = "Status";
            dcStatus.ColumnName = "Status";
            dcStatus.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcStatus);

            DataColumn dcLink = new DataColumn();
            dcLink.DataType = System.Type.GetType("System.String");
            dcLink.AllowDBNull = true;
            dcLink.Caption = "Link";
            dcLink.ColumnName = "Link";
            dcLink.DefaultValue = "";
            // Add the column to the table. 
            table.Columns.Add(dcLink);

            #endregion

            DataRow row;
            foreach (Task t in taskList)
            {
                String openItemUrl = ItemUrlProtocol + service.Url.Host + ItemUrlFolder + t.WebClientReadFormQueryString;

                row = table.NewRow();
                row["Subject"] = cleanHighAsciiForXML(t.Subject);
                row["Created"] = t.DateTimeCreated.ToShortDateString() + " " + t.DateTimeCreated.ToShortTimeString();
                row["StartDate"] = t.StartDate.Value.ToShortDateString() + " " + t.StartDate.Value.ToShortTimeString();
                row["DueDate"] = t.DueDate.Value.ToShortDateString() + " " + t.DueDate.Value.ToShortTimeString();
                row["Importance"] = t.Importance.ToString();
                row["Status"] = t.Status.ToString();
                row["Link"] = openItemUrl;

                table.Rows.Add(row);
            }

            rsp.TableData = table;
        }
        catch (Exception e)
        {
            if (!(int.TryParse(e.Message, out rsp.StatusCode)))
            {
                //if the parsing fails, then set a default value of 500
                rsp.StatusCode = 500;
            }
            rsp.Msg = e.Message;

            success = false;
            miscData = "{ \"ErrMsg\":\"" + rsp.Msg + "\" }";
            rsp.StackTrace = e.StackTrace;
            if (e.InnerException != null) { rsp.InnerException = e.InnerException.ToString(); }
        }

        //audit
        addAuditLog(apiKey, emlUserAddress, "GetTasks", success, miscData);

        return rsp;
    }

    #endregion
}