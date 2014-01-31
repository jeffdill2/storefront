using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Security.Cryptography;
using System.Configuration;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Text;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace StorefrontCreditCardSite
{
    public partial class StorefrontPayment : System.Web.UI.Page
    {
        #region "    Properties...    "

        //////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////// DECLARING PROGRAM PROPERTIES
        //////////////////////////////////////////////////////////////////////////////////////////

        // Misc. values used for processing
        static int iEmailAttempts = 0;
        
        // Values pulled from the URL
        static string strDirectResponseOrderNumber = "";
        
        // Values needed for Authorize.net
        static string strTotalChargeAmount = "";
        static string strPrintAmount = "";
        static string strProductDescription = "";

        // Values retrieved from the Post Process table
        static string strProjectName = "";
        static string strProductPrice = "";
        static string strQuantity = "";
        static string strPostageAmount = "";
        static string strShippingAmount = "";
        static string strTaxAmount = "";
        static string strOrderFee = "";

        // Retrieving values from the web config file.

        // TESTING VALUES
        static bool bolRunInTestMode = ConfigurationManager.AppSettings["RunInTestMode"].ToUpper().Equals("TRUE") ? true : false;

        // LIVE VALUES
        static string strAPILoginID = ConfigurationManager.AppSettings["APILoginID"];
        static string strTransactionKey = ConfigurationManager.AppSettings["TransactionKey"];
        static string strRelayResponseURL = ConfigurationManager.AppSettings["RelayResponseURL"];
        static string strButtonLabel = ConfigurationManager.AppSettings["ButtonLabel"];
        static string strLogFile = ConfigurationManager.AppSettings["LogFile"];
        static string strSQLPostProcessTable = ConfigurationManager.AppSettings["SQLPostProcessTable"];
        static string strEmailToTestMode = ConfigurationManager.AppSettings["EmailToTestMode"];
        static string strEmailSMTPHost = ConfigurationManager.AppSettings["EmailSMTPHost"];
        static string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
        static string strErrorLogEmailTo = ConfigurationManager.AppSettings["ErrorLogEmailTo"];
        static string strSQLConnection = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;

        #endregion


        #region "    Page Load...    "

        protected void Page_Load(object sender, EventArgs e)
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////// INITIATING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strInvoice = DateTime.Now.ToString("yyyyMMddhhmmss");
            string strTimeStamp = ((int)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds).ToString();
            Random rdmRandomValue = new Random();
            string strSequence = (rdmRandomValue.Next(0, 1000)).ToString();
            strDirectResponseOrderNumber = Request.QueryString["OrderNum"];

            //////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////// RETRIEVING THE ORDER INFORMATION FROM THE POST PROCESS TABLE
            //////////////////////////////////////////////////////////////////////////////////////////
            RetrieveOrder();

            //////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////// PULLING THE CHARGE AMOUNT FROM THE POST PROCESS TABLE
            //////////////////////////////////////////////////////////////////////////////////////////
            if (string.IsNullOrEmpty(strProjectName))
            {
                strProductDescription = " ** ERROR - UNABLE TO QUERY ** ";
                strQuantity = " ** ERROR - UNABLE TO QUERY ** ";
                strProductPrice = " ** ERROR - UNABLE TO QUERY ** ";
                strOrderFee = " ** ERROR - UNABLE TO QUERY ** ";
                strPostageAmount = " ** ERROR - UNABLE TO QUERY ** ";
                strShippingAmount = " ** ERROR - UNABLE TO QUERY ** ";
                strTaxAmount = " ** ERROR - UNABLE TO QUERY ** ";
                strPrintAmount = " ** ERROR - UNABLE TO QUERY ** ";
                strTotalChargeAmount = " ** ERROR - UNABLE TO QUERY ** ";
            }
            else
            {
                strProductDescription = strProjectName;
                strPrintAmount = (decimal.Parse(strQuantity) * decimal.Parse(strProductPrice)).ToString();
                strTotalChargeAmount = (decimal.Parse(strPrintAmount) + decimal.Parse(strPostageAmount) + decimal.Parse(strShippingAmount) + decimal.Parse(strTaxAmount) + decimal.Parse(strOrderFee)).ToString();
            }

            //////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////// PROPERLY FORMATTING THE CURRENCY AMOUNTS WITH TWO DECIMAL PLACES
            //////////////////////////////////////////////////////////////////////////////////////////
            strProductPrice = FormatCurrency(strProductPrice);
            strOrderFee = FormatCurrency(strOrderFee);
            strPrintAmount = FormatCurrency(strPrintAmount);
            strPostageAmount = FormatCurrency(strPostageAmount);
            strShippingAmount = FormatCurrency(strShippingAmount);
            strTaxAmount = FormatCurrency(strTaxAmount);
            strTotalChargeAmount = FormatCurrency(strTotalChargeAmount);

            //////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////////////////// POSTING THE PAYMENT INFORMATION TO THE FORM
            //////////////////////////////////////////////////////////////////////////////////////////

            // Creating the fingerprint.
            string strFingerprint = HMAC_MD5(strTransactionKey, strAPILoginID + "^" + strSequence + "^" + strTimeStamp + "^" + strTotalChargeAmount + "^");
            
            // Updating the page with the correct product description and amounts.
            productDescriptionSpan.InnerHtml = strProductDescription;
            printQuantitySpan.InnerHtml = strQuantity;
            perPiecePriceSpan.InnerHtml = strProductPrice;
            printAmountSpan.InnerHtml = strPrintAmount;
            postageAmountSpan.InnerHtml = strPostageAmount;
            shippingAmountSpan.InnerHtml = strShippingAmount;
            taxAmountSpan.InnerHtml = strTaxAmount;
            orderFeeSpan.InnerHtml = strOrderFee;
            totalChargeAmountSpan.InnerHtml = strTotalChargeAmount;

            // Update the fields in the actual form.
            x_login.Value = strAPILoginID;
            x_amount.Value = strTotalChargeAmount;
            x_description.Value = strProductDescription;
            buttonLabel.Value = strButtonLabel;
            x_test_request.Value = (bolRunInTestMode ? "true" : "false");
            x_invoice_num.Value = strInvoice;
            x_fp_sequence.Value = strSequence;
            x_fp_timestamp.Value = strTimeStamp;
            x_fp_hash.Value = strFingerprint;
            x_invoice_num.Value = strDirectResponseOrderNumber;
            x_relay_response.Value = "true";
            x_relay_url.Value = strRelayResponseURL;
            x_relay_always.Value = "true";
        }

        #endregion


        #region "    Properly Format Currency...    "

        static string FormatCurrency(string strInputValue)
        {
            string strFormattedValue = "";
            
            if (!strInputValue.Contains("."))
            {
                strFormattedValue = strInputValue + ".00";
            }
            else
            {
                if (!(strInputValue.Substring(strInputValue.Length - 3, 1) == "."))
                {
                    if (strInputValue.Substring(strInputValue.Length - 2, 1) == ".")
                    {
                        strFormattedValue = strInputValue + "0";
                    }
                    else
                    {
                        if (strInputValue.Substring(strInputValue.Length - 1, 1) == "0")
                        {
                            strFormattedValue = strInputValue.Substring(0, strInputValue.Length - 1);
                        }
                    }
                }
            }

            return strFormattedValue;
        }

        #endregion


        #region "    Retrieve Order...    "

        static bool RetrieveOrder()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strSQLCommand = "SELECT * FROM " + strSQLPostProcessTable + " WHERE Direct_Response_Order_Number = '" + strDirectResponseOrderNumber + "'";

            SqlConnection sqlDBConnection = new SqlConnection(strSQLConnection);
            SqlCommand sqlDBCommand = new SqlCommand(strSQLCommand, sqlDBConnection);
            SqlDataAdapter sqlDBDataAdapter = new SqlDataAdapter(sqlDBCommand);
            DataSet dataSQLTable = new DataSet();

            try
            {
                sqlDBDataAdapter.Fill(dataSQLTable);

                foreach (DataRow dataSQLTableRow in dataSQLTable.Tables[0].Rows)
                {
                    strProjectName = dataSQLTableRow["Project_Name"].ToString();
                    strProductPrice = dataSQLTableRow["Product_Price"].ToString().Trim();
                    strOrderFee = dataSQLTableRow["Order_Fee"].ToString().Trim();
                    strQuantity = dataSQLTableRow["Quantity"].ToString().Trim();
                    strPostageAmount = dataSQLTableRow["Postage_Amount"].ToString().Trim();
                    strShippingAmount = dataSQLTableRow["Shipping_Amount"].ToString().Trim();
                    strTaxAmount = dataSQLTableRow["Tax_Amount"].ToString().Trim();
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
            string strEmailAttachment = "";
            SmtpClient smtpSendEmail = new SmtpClient(strEmailSMTPHost);

            try
            {
                switch (strEmailType)
                {
                    case "ERROR":
                        strFinalEmail = strErrorLogEmailTo;

                        strEmailSubject = "Storefront Credit Card Site - ERROR HAS OCCURED";
                        strEmailBody = "<body style=\"font-family: Arial; font-size: 12px\">" +
                                       "Please review attached log file for details." +
                                       "<br/><br/>" +
                                       "Direct Response Order Number :    " + strDirectResponseOrderNumber +
                                       "</body>";

                        strEmailAttachment = strLogFile;

                        break;

                    default:
                        break;
                }

                if (bolRunInTestMode)
                {
                    strFinalEmail = strEmailToTestMode;
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
                    msgSendEmail.From = new MailAddress(strEmailFrom);
                    msgSendEmail.Subject = strEmailSubject;
                    msgSendEmail.Body = strEmailBody;

                    if (!string.IsNullOrEmpty(strEmailAttachment) && File.Exists(strEmailAttachment))
                    {
                        using (Attachment emailAttachment = new Attachment(strEmailAttachment))
                        {
                            msgSendEmail.Attachments.Add(emailAttachment);

                            smtpSendEmail.Send(msgSendEmail);
                        }
                    }
                    else
                    {
                        smtpSendEmail.Send(msgSendEmail);
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


        #region "    HMAC MD5 Authorize.net Wrapper...    "

        string HMAC_MD5(string key, string value)
        {
            // This is a wrapper for the VB.NET's built-in HMACMD5 functionality
            // This function takes the data and key as strings and returns the hash as a hexadecimal value

            // The first two lines take the input values and convert them from strings to Byte arrays
            byte[] HMACkey = (new System.Text.ASCIIEncoding()).GetBytes(key);
            byte[] HMACdata = (new System.Text.ASCIIEncoding()).GetBytes(value);

            // create a HMACMD5 object with the key set
            HMACMD5 myhmacMD5 = new HMACMD5(HMACkey);

            //calculate the hash (returns a byte array)
            byte[] HMAChash = myhmacMD5.ComputeHash(HMACdata);

            //loop through the byte array and add append each piece to a string to obtain a hash string
            string fingerprint = "";
            for (int i = 0; i < HMAChash.Length; i++)
            {
                fingerprint += HMAChash[i].ToString("x").PadLeft(2, '0');
            }

            return fingerprint;
        }

        #endregion
    }
}
