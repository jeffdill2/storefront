<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PaymentResponse.aspx.cs" Inherits="StorefrontCreditCardSite.PaymentResponse"%>

<%@ Import Namespace="System.Web"%>
<%@ Import Namespace="System.Web.UI"%>
<%@ Import Namespace="System.Data"%>
<%@ Import Namespace="System.Data.Sql"%>
<%@ Import Namespace="System.Data.SqlClient"%>
<%@ Import Namespace="System.IO"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%
Response.Expires = 0;
Response.Buffer = true;
Response.Flush();

string strResponseCode = "";
string strResponseReasonText = "";
string strResponseReasonCode = "";
string strAVSCode = "";
string strTransID = "";
string strAmount = "";
string strAuthCode = "";
string strDirectResponseOrderNumber = "";
string strReceiptLink = "";

// Retrieving and defining Form Data from Post command body from Authorize.Net
if (!string.IsNullOrEmpty(Request.Form["x_response_code"]))
{
    strResponseCode = Request.Form["x_response_code"];
}

if (!string.IsNullOrEmpty(Request.Form["x_response_reason_text"]))
{
    strResponseReasonText = Request.Form["x_response_reason_text"];
}

if (!string.IsNullOrEmpty(Request.Form["x_response_reason_code"]))
{
    strResponseReasonCode = Request.Form["x_response_reason_code"];
}

if (!string.IsNullOrEmpty(Request.Form["x_avs_code"]))
{
    strAVSCode = Request.Form["x_avs_code"];
}

if (!string.IsNullOrEmpty(Request.Form["x_Trans_ID"]))
{
    strTransID = Request.Form["x_Trans_ID"];
}

if (!string.IsNullOrEmpty(Request.Form["x_Auth_Code"]))
{
    strAuthCode = Request.Form["x_Auth_Code"];
}

if (!string.IsNullOrEmpty(Request.Form["x_Amount"]))
{
    strAmount = Request.Form["x_Amount"];
}

if (!string.IsNullOrEmpty(Request.Form["x_invoice_num"]))
{
    strDirectResponseOrderNumber = Request.Form["x_invoice_num"];
}
    
strReceiptLink = "http://www.authorizenet.com";

//Print a receipt page
%>

<html>
    <head id="Head1" runat="server">
        <title>Transaction Receipt Page</title>
        
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
    </head>
    
    <body bgcolor="#FFFFFF">
<%
        // Test to see if this is a test transaction.
        if (strTransID == "0" && strResponseCode == "1")
        {        
            // If so, print it to the screen, so we know that the transaction will not be processed.
%>
            <table align="center">
	            <tr>
	                <th>
	                    <font size="5" color="red" face="arial">
	                        TEST MODE
	                    </font>
	                </th>
    	            
			        <th valign="top">
			            <font size="1" color="black" face="arial">
			                This transaction will <b><u>NOT</u></b> be processed because your account is in Test Mode.
			            </font>
			        </th>
	            </tr>
            </table>
<%
        }
%>
	    <br/>
	    <br/>
<%
        // Test to see if the transaction resulted in Approvavl, Decline or Error
	    switch (strResponseCode)
        {
            case "1":
%>
                <table align="center">
		            <tr>
			            <th>
                            <font size="3" color="#000000" face="Verdana, Arial, Helvetica, sans-serif">
                                This transaction has been approved.
                            </font>
                        </th>
	                </tr>
	            </table>
<%              
                // Config file variables
                string strLogFile = ConfigurationManager.AppSettings["LogFile"];
                string strSQLPostProcessTable = ConfigurationManager.AppSettings["SQLPostProcessTable"];
                string strSQLConnection = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;
                
                // SQL variables
                SqlConnection sqlDBConnection = new SqlConnection(strSQLConnection);
                string strSQLUpdateCommand = "UPDATE " + strSQLPostProcessTable + " SET Job_Status = 'READY TO PROCESS' WHERE Direct_Response_Order_Number = '" + strDirectResponseOrderNumber + "'";
                SqlCommand sqlDBUpdateCommand = new SqlCommand(strSQLUpdateCommand, sqlDBConnection);

                try
                {
                    sqlDBConnection.Open();
                    sqlDBUpdateCommand.ExecuteNonQuery();
                }
                catch (Exception exception)
                {
                    //try
                    //{
                    //    File.AppendAllText(strLogFile, DateTime.Now.ToString() + " :  " + exception.ToString() + Environment.NewLine + Environment.NewLine + "-----------------------" + Environment.NewLine + Environment.NewLine);
                    //}
                    //catch
                    //{
                    //}
                }
                finally
                {
                    sqlDBConnection.Close();
                    sqlDBConnection.Dispose();
                }
                
                break;    

            case "2":
%>
                <table align="center">
	                <tr>
		                <th width="312">
                            <font size="3" color="#000000" face="Verdana, Arial, Helvetica, sans-serif">
                                This transaction has been declined.
                            </font>
                        </th>
	                </tr>
	            </table>
<%
                break;
                        
            default:
%>
                <table align="center">
	                <tr>
		                <th colspan="2">
		                    <font size="3" color="Red" face="Verdana, Arial, Helvetica, sans-serif">
		                        There was an error processing this transaction.
		                    </font>
		                </th>
	                </tr>
	            </table>
<%
                break;
        }
%>
	    <br/>
	    <br/>
	    
	    <table align="center" width="60%">
	        <tr>	
		        <td align="right" width="175" valign="top">
		            <font size="2" color="black" face="arial">
		                <b>Amount:</b>
		            </font>
		        </td>
		        
		        <td align="left">
		            <font size="2" color="black" face="arial">
		                $<%= strAmount%>
		            </font>
		        </td>
	        </tr>
	
	        <tr>
		        <td align="right" width="175" valign="top">
		            <font size="2" color="black" face="arial">
		                <b>Transaction ID:</b>
		            </font>
		        </td>
		        
		        <td align="left">
		            <font size="2" color="black" face="arial">
<%
                        switch (strTransID)
                        {
                            case "0":
                                Response.Write("Not Applicable.");
                                break;
                                
                            default:
                                Response.Write(strTransID);
                                break;
                        }
%>
                    </font>
	            </td>
	        </tr>

	        <tr>
		        <td align="right" width="175" valign="top">
		            <font size="2" color="black" face="arial">
		                <b>Authorization Code:</b>
		            </font>
		        </td>
		        
		        <td align="left">
		            <font size="2" color="black" face="arial">
<% 
                        switch (strAuthCode)
                        {
                            case "000000":
                                Response.Write("Not Applicable.");
                                break;
                                
                            default:
                                Response.Write(strAuthCode);
                                break;
                        }
%>
                    </font>
		        </td>
		    </tr>
	        
	        <tr>
		        <td align="right" width="175" valign="top">
		            <font size="2" color="black" face="arial">
		                <b>Response Reason:</b>
		            </font>
		        </td>
		        
		        <td align="left">
		            <font size="2" color="black" face="arial">
		                (<%= strResponseReasonCode%>)&nbsp;<%= strResponseReasonText%>
		            </font>
		            
		            <font size="1" color="black" face="arial"/>
		        </td>
		    </tr>
	        
	        <tr>
		        <td align="right" width="175" valign="top">
		            <font size="2" color="black" face="arial">
		                <b>Address Verification:</b>
		            </font>
		        </td>
		        
		        <td align="left">
		            <font size="2" color="black" face="arial">
<%
                        // Turn the AVS code into the corresponding text string.
                        switch (strAVSCode)
                        {
                            case "A":
                                Response.Write("Address (Street) matches, ZIP does not.");
                                break;
                                
                            case "B":
                                Response.Write("Address Information Not Provided for AVS Check.");
                                break;
                                
                            case "C":
                                Response.Write("Street address and Postal Code not verified for international transaction due to incompatible formats. (Acquirer sent both street address and Postal Code.)");
                                break;
                                
                            case "D":
                                Response.Write("International Transaction:  Street address and Postal Code match.");
                                break;
                                
                            case "E":
                                Response.Write("AVS Error.");
                                break;
                                
                            case "G":
                                Response.Write("Non U.S. Card Issuing Bank.");
                                break;
                                
                            case "N":
                                Response.Write("No Match on Address (Street) or ZIP.");
                                break;
                                
                            case "P":
                                Response.Write("AVS not applicable for this transaction.");
                                break;
                                
                            case "R":
                                Response.Write("Retry. System unavailable or timed out.");
                                break;
                                
                            case "S":
                                Response.Write("Service not supported by issuer.");
                                break;
                                
                            case "U":
                                Response.Write("Address information is unavailable.");
                                break;
                                
                            case "W":
                                Response.Write("9 digit ZIP matches, Address (Street) does not.");
                                break;
                                
                            case "X":
                                Response.Write("Address (Street) and 9 digit ZIP match.");
                                break;
                            
                            case "Y":
                                Response.Write("Address (Street) and 5 digit ZIP match.");
                                break;
                                
                            case "Z":
                                Response.Write("5 digit ZIP matches, Address (Street) does not.");
                                break;
                                
                            default:
                                Response.Write("The address verification system returned an unknown value.");
                                break;
                        }
%>
                    </font>
		        </td>
	        </tr>
	    </table>
	</body>
</html>

<%
Response.Flush();
%>