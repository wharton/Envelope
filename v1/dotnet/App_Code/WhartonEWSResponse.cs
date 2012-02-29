using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

/// <summary>
/// Summary description for WhartonEWSResponse
/// </summary>
public class WhartonEWSResponse
{
	public WhartonEWSResponse()
	{
        StatusCode = 200;
        Msg = "OK";
	}

    public int StatusCode = 200;
    public String Msg;
    public Object SimpleData;
    public DataTable TableData;
    public String StackTrace;
    public String InnerException;
}
