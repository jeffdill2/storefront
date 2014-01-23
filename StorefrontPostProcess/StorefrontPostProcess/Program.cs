using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;

namespace StorefrontPostProcess
{
    class Program
    {
        #region "    Properties...    "

        //////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////// DECLARING PROGRAM PROPERTIES
        //////////////////////////////////////////////////////////////////////////////////////////

        // Values used to drive how program runs.
        static bool bolEmailSuccess = false;
        static int iEmailAttempts = 0;

        // Values retrieved from the Post Process table.
        static string strDirectResponseOrderNumber = "";
        static string strProjectName = "";
        static string strProductID = "";
        static double dblProductPrice = 0.0;
        static double dblOrderFee = 0.0;
        static string strJobStatus = "";
        static int iQuantity = 0;
        static double dblPostageAmount = 0.0;
        static string strCreditCardRequired = "";
        static string strJobType = "";

        // Values retrieved from the Log table.
        static string strInternalEmailCC = "";
        static string strStorefrontOrderID = "";
        static string strMailingPrintOutput = "";
        static string strSortedData = "";
        static string strProductionPDFFile = "";
        static bool bolPassToApogee = false;
        static DateTime dtEndTime;
        static string strWatermarkPDFFile = "";
        static string strDeliveryEmail = "";
        static string strUserEmail = "";
        static string strUserName = "";
        static string strMailDate = "";
        static string strDeliveryDate = "";
        static string strAdditionalNotes = "";

        // Values retrieved from the values in the Log table.
        static string strPresortDataJobFolder = "";
        static string strOutputJobFolder = "";

        // Retrieving values from the App config file.

        //TESTING VALUES
        static bool bolRunInTestMode = ConfigurationManager.AppSettings["RunInTestMode"].ToUpper().Equals("TRUE") ? true : false;
        static string strEmailToTestMode = ConfigurationManager.AppSettings["EmailToTestMode"];
        static string strEmailCCTestMode = ConfigurationManager.AppSettings["EmailCCTestMode"];

        // LIVE VALUES
        static string strEmailSMTPHost = ConfigurationManager.AppSettings["EmailSMTPHost"];
        static string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
        static string strEmailCC = ConfigurationManager.AppSettings["EmailCC"];
        static string strErrorLogEmailTo = ConfigurationManager.AppSettings["ErrorLogEmailTo"];
        static string strLogFile = ConfigurationManager.AppSettings["LogFile"];
        static string strInternalEmailTo = ConfigurationManager.AppSettings["InternalEmailTo"];
        static string strSQLPostProcessTable = ConfigurationManager.AppSettings["SQLPostProcessTable"];
        static string strSQLLogTable = ConfigurationManager.AppSettings["SQLLogTable"];
        static string strDirectResponseShipmentURL = ConfigurationManager.AppSettings["DirectResponseShipmentURL"];
        static string strCreditCardPaymentURL = ConfigurationManager.AppSettings["CreditCardPaymentURL"];
        static string strSQLConnection = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;

        #endregion

        
        #region "    Main Process...    "
        
        static void Main(string[] args)
        {
            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////// PROCESSING ANY ORDERS THAT NEED CREDIT CARD PAYMENT
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!PaymentNeededOrders()) Environment.Exit(1);

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////// PROCESSING ANY ORDERS THAT ARE READY FOR PRODUCTION
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!ReadyToProcessOrders()) Environment.Exit(1);
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                Environment.Exit(2);
            }

            Environment.Exit(0);
        }

        #endregion


        #region "    Payment Needed Orders...    "

        static bool PaymentNeededOrders()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strSQLCommand = "SELECT * FROM " + strSQLPostProcessTable + " WHERE Job_Status = 'PAYMENT NEEDED' AND Email_Sent = 'N'";
            string strSQLOrderCommand;

            SqlConnection sqlDBConnection = new SqlConnection(strSQLConnection);
            SqlCommand sqlDBCommand = new SqlCommand(strSQLCommand, sqlDBConnection);
            SqlCommand sqlDBOrderCommand;
            SqlDataAdapter sqlDBDataAdapter = new SqlDataAdapter(sqlDBCommand);
            SqlDataAdapter sqlDBOrderDataAdapter = new SqlDataAdapter();
            DataTable dataSQLTable = new DataTable();
            DataTable dataSQLOrderTable = new DataTable();

            try
            {
                sqlDBDataAdapter.Fill(dataSQLTable);

                foreach (DataRow dataSQLTableRow in dataSQLTable.Rows)
                {
                    strDirectResponseOrderNumber = dataSQLTableRow["Direct_Response_Order_Number"].ToString().Trim();
                    strProjectName = dataSQLTableRow["Project_Name"].ToString();
                    strProductID = dataSQLTableRow["Product_ID"].ToString().Trim().ToUpper();
                    dblProductPrice = double.Parse(dataSQLTableRow["Product_Price"].ToString().Trim());
                    dblOrderFee = double.Parse(dataSQLTableRow["Order_Fee"].ToString().Trim());
                    strJobStatus = dataSQLTableRow["Job_Status"].ToString().Trim();
                    iQuantity = int.Parse(dataSQLTableRow["Quantity"].ToString().Trim());
                    dblPostageAmount = double.Parse(dataSQLTableRow["Postage_Amount"].ToString().Trim());
                    strCreditCardRequired = dataSQLTableRow["Credit_Card_Required"].ToString().Trim().ToUpper();
                    strJobType = dataSQLTableRow["Job_Type"].ToString().Trim();

                    try
                    {
                        // Setting up our SQL processes.
                        strSQLOrderCommand = "SELECT * FROM " + strSQLLogTable + " WHERE Direct_Response_Order_Number = '" + strDirectResponseOrderNumber + "'";
                        sqlDBOrderCommand = new SqlCommand(strSQLOrderCommand, sqlDBConnection);
                        sqlDBOrderDataAdapter.SelectCommand = sqlDBOrderCommand;
                        sqlDBOrderDataAdapter.Fill(dataSQLOrderTable);

                        // Verifying that there's only one matching record in the Log table.
                        if (dataSQLOrderTable.Rows.Count == 0)
                        {
                            LogFile("A Direct Response Order Number match could not be found in the log table.", true);
                            return false;
                        }
                        else
                        {
                            if (dataSQLOrderTable.Rows.Count > 1)
                            {
                                LogFile("Multiple Direct Response Order Number matches were found in the log table.", true);
                                return false;
                            }
                        }

                        foreach (DataRow dataSQLOrderTableRow in dataSQLOrderTable.Rows)
                        {
                            strInternalEmailCC = dataSQLOrderTableRow["Internal_Email_CC"].ToString().Trim();
                            strDeliveryEmail = dataSQLOrderTableRow["Delivery_Email"].ToString().Trim();
                            strUserEmail = dataSQLOrderTableRow["User_Email"].ToString().Trim();
                            strUserName = dataSQLOrderTableRow["User_Name"].ToString().Trim();
                        }

                        // Clearing the table of it's current contents.
                        dataSQLOrderTable.Clear();
                    }
                    catch (Exception exception)
                    {
                        LogFile(exception.ToString(), true);
                        return false;
                    }

                    // Sending the payment email - if successful the Post Process table will be updated as Email_Sent = Y
                    string strSQLUpdateCommand = "UPDATE " + strSQLPostProcessTable + " SET Email_Sent = 'Y' WHERE Direct_Response_Order_Number = '" + strDirectResponseOrderNumber + "'";
                    SqlCommand sqlDBUpdateCommand = new SqlCommand(strSQLUpdateCommand, sqlDBConnection);

                    if (SendEmail("PAYMENT_NEEDED"))
                    {
                        if (sqlDBConnection.State == ConnectionState.Closed)
                        {
                            sqlDBConnection.Open();
                        }
                        
                        sqlDBUpdateCommand.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }
            finally
            {
                sqlDBConnection.Close();
                sqlDBConnection.Dispose();
                sqlDBCommand.Dispose();
                sqlDBDataAdapter.Dispose();
                dataSQLTable.Dispose();
            }

            return true;
        }

        #endregion


        #region "    Ready To Process Orders...    "

        static bool ReadyToProcessOrders()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strSQLCommand = "SELECT * FROM " + strSQLPostProcessTable + " WHERE Job_Status = 'READY TO PROCESS'";
            string strSQLOrderCommand;
            string strSQLUpdateCommand;

            SqlConnection sqlDBConnection = new SqlConnection(strSQLConnection);
            SqlCommand sqlDBCommand = new SqlCommand(strSQLCommand, sqlDBConnection);
            SqlCommand sqlDBOrderCommand;
            SqlCommand sqlDBUpdateCommand;
            SqlDataAdapter sqlDBDataAdapter = new SqlDataAdapter(sqlDBCommand);
            SqlDataAdapter sqlDBOrderDataAdapter = new SqlDataAdapter();
            DataTable dataSQLTable = new DataTable();
            DataTable dataSQLOrderTable = new DataTable();
            
            try
            {
                sqlDBDataAdapter.Fill(dataSQLTable);

                foreach (DataRow dataSQLTableRow in dataSQLTable.Rows)
                {
                    strDirectResponseOrderNumber = dataSQLTableRow["Direct_Response_Order_Number"].ToString().Trim();
                    strProjectName = dataSQLTableRow["Project_Name"].ToString();
                    strProductID = dataSQLTableRow["Product_ID"].ToString().Trim().ToUpper();
                    dblProductPrice = double.Parse(dataSQLTableRow["Product_Price"].ToString().Trim());
                    strJobStatus = dataSQLTableRow["Job_Status"].ToString().Trim();
                    iQuantity = int.Parse(dataSQLTableRow["Quantity"].ToString().Trim());
                    dblPostageAmount = double.Parse(dataSQLTableRow["Postage_Amount"].ToString().Trim());
                    strCreditCardRequired = dataSQLTableRow["Credit_Card_Required"].ToString().Trim().ToUpper();
                    strJobType = dataSQLTableRow["Job_Type"].ToString().Trim();

                    try
                    {
                        // Setting up our SQL processes.
                        strSQLOrderCommand = "SELECT * FROM " + strSQLLogTable + " WHERE Direct_Response_Order_Number = '" + strDirectResponseOrderNumber + "'";
                        sqlDBOrderCommand = new SqlCommand(strSQLOrderCommand, sqlDBConnection);

                        strSQLUpdateCommand = "UPDATE " + strSQLPostProcessTable + " SET Job_Status = 'PROCESSED' WHERE Direct_Response_Order_Number = '" + strDirectResponseOrderNumber + "'";
                        sqlDBUpdateCommand = new SqlCommand(strSQLUpdateCommand, sqlDBConnection);

                        sqlDBOrderDataAdapter.SelectCommand = sqlDBOrderCommand;
                        sqlDBOrderDataAdapter.Fill(dataSQLOrderTable);

                        // Verifying that there's only one matching record in the Log table.
                        if (dataSQLOrderTable.Rows.Count == 0)
                        {
                            LogFile("A Direct Response Order Number match could not be found in the log table.", true);
                            return false;
                        }
                        else
                        {
                            if (dataSQLOrderTable.Rows.Count > 1)
                            {
                                LogFile("Multiple Direct Response Order Number matches were found in the log table.", true);
                                return false;
                            }
                        }

                        foreach (DataRow dataSQLOrderTableRow in dataSQLOrderTable.Rows)
                        {
                            strInternalEmailCC = dataSQLOrderTableRow["Internal_Email_CC"].ToString().Trim();
                            strStorefrontOrderID = dataSQLOrderTableRow["Storefront_Order_ID"].ToString().Trim();
                            strMailingPrintOutput = dataSQLOrderTableRow["Mailing_Print_Output"].ToString().Trim();
                            strSortedData = dataSQLOrderTableRow["Sorted_Data"].ToString().Trim();
                            strProductionPDFFile = dataSQLOrderTableRow["Production_PDF_File"].ToString().Trim();
                            bolPassToApogee = dataSQLOrderTableRow["Pass_To_Apogee"].ToString().Trim().ToUpper().Equals("Y") ? true : false;
                            dtEndTime = DateTime.Parse(dataSQLOrderTableRow["End_Time"].ToString());
                            strWatermarkPDFFile = dataSQLOrderTableRow["Watermark_PDF_File"].ToString().Trim();
                            strDeliveryEmail = dataSQLOrderTableRow["Delivery_Email"].ToString().Trim();
                            strUserEmail = dataSQLOrderTableRow["User_Email"].ToString().Trim();
                            strUserName = dataSQLOrderTableRow["User_Name"].ToString().Trim();
                            strMailDate = dataSQLOrderTableRow["Mail_Date"].ToString().Trim();
                            strDeliveryDate = dataSQLOrderTableRow["Delivery_Date"].ToString().Trim();
                            strAdditionalNotes = dataSQLOrderTableRow["Additional_Notes"].ToString().Trim();
                        }

                        // Setting some values based on file names.
                        if (!string.IsNullOrEmpty(strSortedData))
                        {
                            strPresortDataJobFolder = Path.GetDirectoryName(strSortedData);
                        }

                        if (!string.IsNullOrEmpty(strProductionPDFFile))
                        {
                            strOutputJobFolder = Path.GetDirectoryName(strProductionPDFFile);
                        }

                        // Clearing the table of it's current contents.
                        dataSQLOrderTable.Clear();
                    }
                    catch (Exception exception)
                    {
                        LogFile(exception.ToString(), true);
                        return false;
                    }

                    // Updating the Post Process table to 'PROCESSED' for each processed job.
                    if (sqlDBConnection.State == ConnectionState.Closed)
                    {
                        sqlDBConnection.Open();
                    }

                    sqlDBUpdateCommand.ExecuteNonQuery();

                    // Sending the internal and client emails to let them know the job has been processed.
                    SendEmail("PROCESSED_INTERNAL");
                    SendEmail("PROCESSED_CLIENT");
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }
            finally
            {
                sqlDBConnection.Close();
                sqlDBConnection.Dispose();
                sqlDBCommand.Dispose();
                sqlDBDataAdapter.Dispose();
                dataSQLTable.Dispose();
            }

            return true;
        }

        #endregion


        #region "    Send Email...    "

        static bool SendEmail(string strEmailType)
        {
            iEmailAttempts++;
            
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strEmailSubject = "";
            string strEmailBody = "";
            string strFinalEmail = "";
            string strFinalEmailCC = "";
            string strEmailAttachment = "";

            try
            {
                switch (strEmailType)
                {
                    case "ERROR":
                        strFinalEmail = strErrorLogEmailTo;
                        strFinalEmailCC = strErrorLogEmailTo;

                        strEmailSubject = "Storefront Post Process - ERROR HAS OCCURED";
                        strEmailBody = "<body style=\"font-family: Arial, sans-serif; font-size: 12px\">" +
                                       "Please review attached log file for details." +
                                       "<br/><br/>" +
                                       "Direct Response Order Number :    " + strDirectResponseOrderNumber +
                                       "</body>";

                        strEmailAttachment = strLogFile;

                        break;

                    case "PROCESSED_INTERNAL":
                        strFinalEmail = strInternalEmailTo;
                        strFinalEmailCC = strInternalEmailCC;

                        if (strJobType.Equals("EMAILPDF"))
                        {
                            strEmailSubject = strProjectName + " - JOB PROCESSED";
                        }
                        else
                        {
                            strEmailSubject = strProjectName + " - READY FOR PRODUCTION";
                        }

                        strEmailBody = "<body style=\"font-family: Arial, sans-serif; font-size: 12px\">" +
                                       "Processing details below:" +
                                       "<br/><br/>" +
                                       "<b>Order Number :&nbsp;&nbsp;&nbsp;</b>" + strDirectResponseOrderNumber + "<br/>" +
                                       "<b>Storefront ID Number :&nbsp;&nbsp;&nbsp;</b>" + strStorefrontOrderID + "<br/>" +
                                       "<b>Output Type :&nbsp;&nbsp;&nbsp;</b>" + strJobType + "<br/>" +
                                       "<b>Quantity :&nbsp;&nbsp;&nbsp;</b>" + iQuantity.ToString() + "<br/>" +
                                       "<b>Print File Name :&nbsp;&nbsp;&nbsp;</b>" + (strJobType.Equals("EMAILPDF") ? "N/A" : (strJobType.Equals("PRINTMAIL") && strMailingPrintOutput.Equals("INKJET")) ? Path.GetFileName(strSortedData) : Path.GetFileName(strProductionPDFFile)) + "<br/>" +
                                       "<b>Print File Location :&nbsp;&nbsp;&nbsp;</b>" + (strJobType.Equals("EMAILPDF") || (strJobType.Equals("PRINTSHIP") && bolPassToApogee) ? "N/A" : (strJobType.Equals("PRINTMAIL") && strMailingPrintOutput.Equals("INKJET")) ? "<a href=\"" + strPresortDataJobFolder.Replace(" ", "%20") + "\">" + strPresortDataJobFolder + "</a>" : "<a href=\"" + strOutputJobFolder.Replace(" ", "%20") + "\">" + strOutputJobFolder + "</a>") + "<br/>" +
                                       "<b>Mail Date :&nbsp;&nbsp;&nbsp;</b>" + strMailDate + "<br/>" +
                                       "<b>Delivery Date:&nbsp;&nbsp;&nbsp;</b>" + strDeliveryDate + "<br/>" +
                                       "<b>Date/Time of Processing :&nbsp;&nbsp;&nbsp;</b>" + dtEndTime.ToString() +
                                       "<br/><br/>" +
                                       "<b><u>Additional Notes:</u></b><br/>" +
                                       (string.IsNullOrEmpty(strAdditionalNotes) ? "N/A" : strAdditionalNotes) +
                                       "</body>";

                        strEmailAttachment = strWatermarkPDFFile;

                        break;

                    case "PROCESSED_CLIENT":
                        strFinalEmail = (strDeliveryEmail.Equals("")) ? strUserEmail : strDeliveryEmail;
                        strFinalEmailCC = strEmailCC;

                        strEmailSubject = strProjectName;

                        if (strJobType.Equals("EMAILPDF"))
                        {
                            strEmailBody = "<body style=\"font-family: Arial, sans-serif; font-size: 12px\">" +
                                           "Hello " + strUserName + "," +
                                           "<br/><br/>" +
                                           "The following job has been processed:<br/>" +
                                           "<b>" + strProjectName + "</b>" +
                                           "<br/><br/>" +
                                           "Please see attached PDF file." +
                                           "<br/><br/>" +
                                           "Thanks!<br/>" +
                                           "Heeter" +
                                           "</body>";

                            strEmailAttachment = strProductionPDFFile;
                        }
                        else
                        {
                            strEmailBody = "<body style=\"font-family: Arial, sans-serif; font-size: 12px\">" +
                                           "Hello " + strUserName + "," +
                                           "<br/><br/>" +
                                           "The following job has been processed and moved into Production :<br/>" +
                                           "<b>" + strProjectName + "</b>" +
                                           "<br/><br/>" +
                                           "Below is the link to your shipping order :<br/>" +
                                           "<a href=\"" + strDirectResponseShipmentURL + strDirectResponseOrderNumber + "\">" + strDirectResponseShipmentURL + strDirectResponseOrderNumber + "</a>" +
                                           "<br/><br/>" +
                                           "Thanks!<br/>" +
                                           "Heeter" +
                                           "</body>";

                            strEmailAttachment = strWatermarkPDFFile;
                        }

                        break;

                    case "PAYMENT_NEEDED":
                        strFinalEmail = (strDeliveryEmail.Equals("")) ? strUserEmail : strDeliveryEmail;
                        strFinalEmailCC = strEmailCC;

                        strEmailSubject = strProjectName + " - PAYMENT NEEDED";

                        strEmailBody = "<body style=\"font-family: Arial, sans-serif; font-size: 12px\">" +
                                        "Hello " + strUserName + "," +
                                        "<br/><br/>" +
                                        "The following job is awaiting credit card payment :<br/>" +
                                        "<b>" + strProjectName + "</b>" +
                                        "<br/><br/>" +
                                        "Please follow the link below to pay for this order :<br/>" +
                                        "<a href=\"" + strCreditCardPaymentURL + strDirectResponseOrderNumber + "\">" + strCreditCardPaymentURL + strDirectResponseOrderNumber + "</a>" +
                                        "<br/><br/>" +
                                        "Thanks!<br/>" +
                                        "Heeter" +
                                        "</body>";

                        break;

                    default:
                        break;
                }

                if (bolRunInTestMode)
                {
                    strFinalEmail = strEmailToTestMode;
                    strFinalEmailCC = strEmailCCTestMode;
                }
            }
            catch (Exception exception)
            {
                if (iEmailAttempts <= 1)
                {
                    LogFile(exception.ToString(), true);
                }
                
                return false;
            }


            //////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////// CREATING AND SENDING THE MAIL MESSAGE
            //////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                using (MailMessage msgSendEmail = new MailMessage())
                {
                    msgSendEmail.IsBodyHtml = true;
                    msgSendEmail.To.Add(strFinalEmail);
                    msgSendEmail.CC.Add(strFinalEmailCC);
                    msgSendEmail.From = new MailAddress(strEmailFrom);
                    msgSendEmail.Subject = strEmailSubject;
                    msgSendEmail.Body = strEmailBody;

                    if (!string.IsNullOrEmpty(strEmailAttachment) && File.Exists(strEmailAttachment))
                    {
                        using (Attachment emailAttachment = new Attachment(strEmailAttachment))
                        {
                            msgSendEmail.Attachments.Add(emailAttachment);

                            using (SmtpClient smtpSendEmail = new SmtpClient(strEmailSMTPHost))
                            {
                                smtpSendEmail.Send(msgSendEmail);
                            }
                        }
                    }
                    else
                    {
                        using (SmtpClient smtpSendEmail = new SmtpClient(strEmailSMTPHost))
                        {
                            smtpSendEmail.Send(msgSendEmail);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                if (iEmailAttempts <= 1)
                {
                    LogFile(exception.ToString(), true);
                }

                return false;
            }      

            return true;
        }

        #endregion


        #region "    Log File...    "

        static void LogFile(string strLogMessage, bool bolSendEmail)
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////// LOGGING REQUESTED VALUES TO THE LOG FILE
            //////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                File.AppendAllText(strLogFile, DateTime.Now.ToString() + " :  " + strLogMessage + Environment.NewLine + Environment.NewLine + "-----------------------" + Environment.NewLine + Environment.NewLine);
            }
            catch (IOException exception)
            {
                LogFile(exception.ToString(), false);
            }

            if (bolSendEmail)
            {
                SendEmail("ERROR");
            }
        }

        #endregion
    }
}
