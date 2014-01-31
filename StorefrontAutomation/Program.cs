using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Net;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Security;
using System.Threading;
using System.Reflection;
using Microsoft.VisualBasic.FileIO;
using Tools;

namespace Storefront_Automation
{
    class Program
    {
        #region "    Properties...    "
        
        //////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////// DECLARING PROGRAM PROPERTIES
        //////////////////////////////////////////////////////////////////////////////////////////
        
        // Values used to drive how program runs.
        static DateTime dtStartTime;
        static DateTime dtEndTime;
        static bool bolPassToApogee = false;
        static int iEmailAttempts = 0;
        static bool bolRegionMismatchThresholdReached = false;
        
        // Values retrieved from interrogating the XML files.
        static string strWFDFullName = "";
        static string strWFDNameOnly = "";
        static string strProjectName = "";
        static string strInputXMLName = "";
        static string strItemDataXMLFullName = "";
        static string strItemDataXMLNameOnly = "";
        static string strInputDataName = "";
        static string strOrderFirstName = "";
        static string strOrderLastName = "";
        static string strLoginID = "";
        static string strCompanyName = "";
        static string strUserEmail = "";
        static string strDeliveryEmail = "";
        static string strMailDate = "";
        static string strDeliveryDate = "";
        static string strOutputType = "";
        static int iInputQuantity = 0;
        static string strShipName = "";
        static string strShipCareOf = "";
        static string strShipAdd1 = "";
        static string strShipAdd2 = "";
        static string strShipCity = "";
        static string strShipState = "";
        static string strShipZip = "";
        static string strRegionCode = "";
        static string strStorefrontOrderID = "";
        static List<string> lstImageFiles = new List<string>();
        static string strProductID = "";
        static double dblProductPrice = 0.0;
        static string strAdditionalNotes = "";
        static string strSplitOverride = "";
        static string strClassOfMail = "";
        static double dblShippingAmount = 0.0;
        static double dblTaxAmount = 0.0;

        // Values retrieved from interrogating project config table.
        static string strDataCasing = "";
        static string strNCOACompany = "";
        static bool bolMultitrac = false;
        static bool bolGenerateClientID = false;
        static string strImportName = "";
        static string strPresortSettings = "";
        static string strFirstClassPresortSettings = "";
        static string strStandardPresortSettings = "";
        static string strDedupeType = "";
        static string strOutputLabel = "";
        static string strMultitracClientID= "";
        static int iDataSplitValue = 0;
        static string strApogeeClientFolder = "";
        static bool bolSendClientEmail = false;
        static string strInternalEmailCC = "";
        static string strPrimaryKitKey = "";
        static string strAdditionalKitKeyField = "";
        static string strMailingPrintOutput = "";
        static string strSimplexOrDuplex = "";
        static string strDROrderPrefix = "";
        static string strDRInventoryName = "";
        static double dblOrderFee = 0.0;
        static bool bolPaymentExpected = false;
        
        // Values used for/retrieved from Pace.
        static string strJobNumber = "";
        static string strClientCode = "";

        // Values used for utilizing web services.
        static string strXMLHeader = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
        static string strSOAPHeader = "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">";

        // Values used for BCC/Multitrac processing.
        static string strMultitracInput = "";
        static string strMultitracOutput = "";
        static string strInvalidsRemoved = "";
        static string strDupesRemoved = "";
        static string strSortedData = "";
        static int iMailingRecords = 0;
        static string strMailingRecords = "";
        static string strMultitracJobNumber = "";
        static string strRegionErrorsRemoved = "";
        static string strBCCListName = "";
        static string strBCCDatabase = "";

        // Job folders needed for processing.
        static string strBCCJobFolder= "";
        static string strInputJobFolder = "";
        static string strPresortDataJobFolder = "";
        static string strWorkingJobFolder = "";
        static string strPNetImagesJobFolder = "";
        //static string strMailDocsJobFolder = "";
        //static string strOutputJobFolder = "";
        static string strDocsJobFolder = "";
        static string strProductionJobFolder = "";
        static string strCSRJobFolder = "";
        static string strJobStatus = "";
        static double dblPostageAmount = 0.0;
        static string strCreditCardRequired = "";

        // Output file values.
        static string strProductionPDFFile = "";
        static string strWatermarkPDFFile = "";
        static string strApogeePDFFile = "";

        // Direct Response values.
        static string strDirectResponseCSV = "";
        static string strDirectResponseOrderNumber = "";
        static string strDirectResponseItemCode = "";

        // Retrieving values from the App config file.

        //TESTING VALUES
        static bool bolRunInTestMode = ConfigurationManager.AppSettings["RunInTestMode"].ToUpper().Equals("TRUE") ? true : false;
        static string strEmailToTestMode = ConfigurationManager.AppSettings["EmailToTestMode"];
        static string strEmailCCTestMode = ConfigurationManager.AppSettings["EmailCCTestMode"];
        static string strMultitracTestClientID = ConfigurationManager.AppSettings["MultitracTestClientID"];
        static string strTestJobNumber = ConfigurationManager.AppSettings["TestJobNumber"];
        static string strTestClientCode = ConfigurationManager.AppSettings["TestClientCode"];
        static string strPDFProofRepository = ConfigurationManager.AppSettings["PDFProofRepository"];

        // LIVE VALUES
        static string strEmailSMTPHost = ConfigurationManager.AppSettings["EmailSMTPHost"];
        static string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
        static string strEmailCC = ConfigurationManager.AppSettings["EmailCC"];
        static string strErrorLogEmailTo = ConfigurationManager.AppSettings["ErrorLogEmailTo"];
        static string strLogFile = ConfigurationManager.AppSettings["LogFile"];
        static string strDecompressedZipFolder = ConfigurationManager.AppSettings["DecompressedZipFolder"];
        static string strStoreFrontFolder = ConfigurationManager.AppSettings["StoreFrontFolder"];
        static string strDataJobsFolder = ConfigurationManager.AppSettings["DataJobsFolder"];
        static string strGeneralJobsFolder = ConfigurationManager.AppSettings["GeneralJobsFolder"];
        static string strProjectConfigFolder = ConfigurationManager.AppSettings["ProjectConfigFolder"];
        static string strGMCConsoleExe = ConfigurationManager.AppSettings["GMCConsoleExe"];
        static string strTaskmasterJobsFolder = ConfigurationManager.AppSettings["TaskmasterJobsFolder"];
        static string strMailDatSettingsFolder = ConfigurationManager.AppSettings["MailDatSettingsFolder"];
        static string strMailDatRepository = ConfigurationManager.AppSettings["MailDatRepository"];
        static string strBCCMailManEXE = ConfigurationManager.AppSettings["BCCMailManEXE"];
        static string strGMCServerIPAddress = ConfigurationManager.AppSettings["GMCServerIPAddress"];
        static string strGMCServerPort = ConfigurationManager.AppSettings["GMCServerPort"];
        static string strMailDatVersion = ConfigurationManager.AppSettings["MailDatVersion"];
        static string strMailDatContactName = ConfigurationManager.AppSettings["MailDatContactName"];
        static string strMailDatContactEmail = ConfigurationManager.AppSettings["MailDatContactEmail"];
        static string strMailDatContactPhone = ConfigurationManager.AppSettings["MailDatContactPhone"];
        static string strMultitracUserName = ConfigurationManager.AppSettings["MultitracUserName"];
        static string strMultitracPassword = ConfigurationManager.AppSettings["MultitracPassword"];
        static string strMultitracEXE = ConfigurationManager.AppSettings["MultitracEXE"];
        static string strDaysToAddForInHome = ConfigurationManager.AppSettings["DaysToAddForInHome"];
        static string strServerUserName = ConfigurationManager.AppSettings["ServerUserName"];
        static string strServerPassword = ConfigurationManager.AppSettings["ServerPassword"];
        static string strHeeterMailerID = ConfigurationManager.AppSettings["HeeterMailerID"];
        static string strZipToRegionTable = ConfigurationManager.AppSettings["ZipToRegionTable"];
        static int iZipToRegionMismatchThreshold = int.Parse(ConfigurationManager.AppSettings["ZipToRegionMismatchThreshold"]);
        static string strHighmarkEmailTo = ConfigurationManager.AppSettings["HighmarkEmailTo"];
        static string strDataInputModule = ConfigurationManager.AppSettings["DataInputModule"];
        static string strParamInputModule = ConfigurationManager.AppSettings["ParamInputModule"];
        static string strDataGeneratorModule = ConfigurationManager.AppSettings["DataGeneratorModule"];
        static string strDigitalOutputModule = ConfigurationManager.AppSettings["DigitalOutputModule"];
        static string strOffsetOutputModule = ConfigurationManager.AppSettings["OffsetOutputModule"];
        static string strWatermarkOutputModule = ConfigurationManager.AppSettings["WatermarkOutputModule"];
        static string strDefaultOutputModule = ConfigurationManager.AppSettings["DefaultOutputModule"];
        static string strImageParamInputModule = ConfigurationManager.AppSettings["ImageParamInputModule"];
        static string strXMLInputModule = ConfigurationManager.AppSettings["XMLInputModule"];
        static string strProductionParameterName = ConfigurationManager.AppSettings["ProductionParameterName"];
        static string strStorefrontIDParameterName = ConfigurationManager.AppSettings["StorefrontIDParameterName"];
        static string strImageParameterRootName = ConfigurationManager.AppSettings["ImageParameterRootName"];
        static string strLowResConfigFile = ConfigurationManager.AppSettings["LowResConfigFile"];
        static string strHighResConfigFile = ConfigurationManager.AppSettings["HighResConfigFile"];
        static string strPDFPrinter = ConfigurationManager.AppSettings["PDFPrinter"];
        static string strTextPrinter = ConfigurationManager.AppSettings["TextPrinter"];
        static string strEnvironmentVariable = ConfigurationManager.AppSettings["EnvironmentVariable"];
        static string strUserInfoTable = ConfigurationManager.AppSettings["UserInfoTable"];
        static string strInternalEmailTo = ConfigurationManager.AppSettings["InternalEmailTo"];
        static string strSQLLogTable = ConfigurationManager.AppSettings["SQLLogTable"];
        static string strSQLProjectConfigTable = ConfigurationManager.AppSettings["SQLProjectConfigTable"];
        static string strSQLPostProcessTable = ConfigurationManager.AppSettings["SQLPostProcessTable"];
        static string strDirectResponseFTP = ConfigurationManager.AppSettings["DirectResponseFTP"];
        static string strDirectResponseUserName = ConfigurationManager.AppSettings["DirectResponseUserName"];
        static string strDirectResponsePassword = ConfigurationManager.AppSettings["DirectResponsePassword"];
        static string strDaysToAddForShipping = ConfigurationManager.AppSettings["DaysToAddForShipping"];
        static string strSQLConnection = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;
        
        #endregion


        #region "    Main Process...    "

        static void Main(string[] aryInputFile)
        {
            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////// CAPTURING THE START TIME OF THE PROGRAM
                //////////////////////////////////////////////////////////////////////////////////////////
                dtStartTime = DateTime.Now;

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////// RETRIEVING THE INPUT XML NAME
                //////////////////////////////////////////////////////////////////////////////////////////
                if (aryInputFile.Length != 0)
                {
                    strInputXMLName = strDecompressedZipFolder + aryInputFile[0];

                    if ((!strInputXMLName.ToLower().Contains("order.")) && (!strInputXMLName.ToLower().Contains(".xml")))
                    {
                        Environment.Exit(1);
                    }
                }
                else
                {
                    LogFile("No XML file was specified as a parameter for the application.", true);
                    Environment.Exit(1);
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////// PARSING XML FILE CREATED FROM STOREFRONT
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!ParseXMLFile()) Environment.Exit(1);

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////// PARSING THE PROJECT CONFIG FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!ParseProjectConfig()) Environment.Exit(1);

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////// CREATING/EDITING THE MAIL/DELIVERY DATE
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!CreateJobDates()) Environment.Exit(1);

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////// MAKING SURE THE WFD EXISTS BEFORE MOVING FORWARD
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!File.Exists(strWFDFullName))
                {
                    LogFile("The WFD - " + strWFDFullName + " - for this project could not be found.", true);

                    Environment.Exit(1);
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                //////// SETTING DIRECT RESPONSE ORDER NUMBER - ALSO USED AS MAIN REFERENCE NUMBER FOR JOB
                //////////////////////////////////////////////////////////////////////////////////////////
                strDirectResponseOrderNumber = strDROrderPrefix + "_" + strStorefrontOrderID;

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////// RUNNING EFI PACE PROCESSES FOR JOB TICKET CREATION **IN DEVELOPMENT**
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!PaceProcess()) Environment.Exit(1);

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////// VERIFYING THAT A DATA FILE EXISTS IF 'PRINTMAIL' OUTPUT TYPE
                //////////////////////////////////////////////////////////////////////////////////////////
                if (strOutputType.Equals("PRINTMAIL"))
                {
                    if (!File.Exists(strInputDataName))
                    {
                        LogFile("The " + strProjectName + " job is identified as a PRINTMAIL job - but no data file can be found.", true);

                        Environment.Exit(1);
                    }
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////// CREATING JOB FOLDER AND MOVING NEEDED FILES INTO PLACE
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!FilePlacement()) Environment.Exit(1);

                //////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////// RUNNING BCC IF 'PRINTMAIL' OUTPUT TYPE
                //////////////////////////////////////////////////////////////////////////////////////////
                if (strOutputType.Equals("PRINTMAIL"))
                {
                    if (!string.IsNullOrEmpty(strRegionCode))
                    {
                        if (!VerifyZipAndRegion()) Environment.Exit(2);
                    }

                    if (!BCCProcess()) Environment.Exit(2);
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////// DETERMINING PAYMENT INFORMATION
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!PaymentInformation()) Environment.Exit(1);

                //////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////// POSTING DIRECT RESPONSE CSV TO FTP
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!DirectResponseProcess()) Environment.Exit(2);

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////////////// RUNNING GMC PROCESSES
                //////////////////////////////////////////////////////////////////////////////////////////
                if (!strOutputType.Equals("PICKANDPACK"))
                {
                    if (!GMCProcess()) Environment.Exit(2);
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////// CAPTURING THE END TIME OF THE PROGRAM
                //////////////////////////////////////////////////////////////////////////////////////////
                dtEndTime = DateTime.Now;

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////// LOGGING ALL PROCESSING INFO INTO SQL TABLES
                //////////////////////////////////////////////////////////////////////////////////////////
                DatabaseLogTable();
                DatabasePostProcessTable();
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                Environment.Exit(3);
            }

            Environment.Exit(0);
        }

        #endregion


        #region "    Parse XML File...    "

        static bool ParseXMLFile()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////// PARSING THE INPUT XML FILE
            //////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                using (XmlTextReader xmlInputXMLStream = new XmlTextReader(strInputXMLName))
                {
                    while (xmlInputXMLStream.Read())
                    {
                        switch (xmlInputXMLStream.Name.ToUpper())
                        {
                            case "PACKAGE":
                                strWFDNameOnly = xmlInputXMLStream.GetAttribute(1);
                                break;

                            case "NAME":
                                string strXMLReadLine;
                                string strFileExtension;

                                strXMLReadLine = xmlInputXMLStream.ReadString().ToString();
                                strFileExtension = Path.GetExtension(strXMLReadLine);
                                strFileExtension = strFileExtension.ToUpper().Replace(".", "");

                                switch (strFileExtension)
                                {
                                    case "XML":
                                        strItemDataXMLNameOnly = strXMLReadLine;
                                        strItemDataXMLFullName = strDecompressedZipFolder + strItemDataXMLNameOnly;
                                        break;

                                    case "CSV":
                                        strInputDataName = strDecompressedZipFolder + strXMLReadLine;
                                        break;

                                    case "JPG": case "JPEG": case "PNG": case "GIF": case "IMG": case "TIFF": case "BMP": case "PDF":
                                        lstImageFiles.Add(strXMLReadLine);
                                        break;

                                    default:
                                        break;
                                }

                                break;

                            case "NETUNITPRICE":
                                dblProductPrice = double.Parse(xmlInputXMLStream.ReadString());
                                break;

                            default:
                                break;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////// PARSING THE ITEM DATA XML FILE
            //////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                using (XmlTextReader xmlItemDataXMLStream = new XmlTextReader(strItemDataXMLFullName))
                {
                    while (xmlItemDataXMLStream.Read())
                    {
                        switch (xmlItemDataXMLStream.Name.ToUpper())
                        {
                            case "FIRSTNAME":
                                strOrderFirstName = xmlItemDataXMLStream.ReadString();
                                break;

                            case "LASTNAME":
                                strOrderLastName = xmlItemDataXMLStream.ReadString();
                                break;

                            case "LOGIN":
                                strLoginID = xmlItemDataXMLStream.ReadString().ToUpper();
                                break;

                            case "EMAIL":
                                strUserEmail = xmlItemDataXMLStream.ReadString();
                                break;

                            case "DELEMAIL":
                                strDeliveryEmail = xmlItemDataXMLStream.ReadString();
                                break;

                            case "MAILDATE":
                                strMailDate = xmlItemDataXMLStream.ReadString();
                                break;

                            /*
                            case "DELDATE":
                                strDeliveryDate = xmlItemDataXMLStream.ReadString();
                                break;
                            */

                            case "OUTPUTTYPE":
                                strOutputType = xmlItemDataXMLStream.ReadString().ToUpper();
                                break;

                            case "QTY":
                                iInputQuantity = int.Parse(xmlItemDataXMLStream.ReadString());
                                break;

                            case "SHIPNAME":
                                strShipName = xmlItemDataXMLStream.ReadString();
                                break;

                            case "SHIPCAREOF":
                                strShipCareOf = xmlItemDataXMLStream.ReadString();
                                break;

                            case "SHIPADD1":
                                strShipAdd1 = xmlItemDataXMLStream.ReadString();
                                break;

                            case "SHIPADD2":
                                strShipAdd2 = xmlItemDataXMLStream.ReadString();
                                break;

                            case "SHIPCITY":
                                strShipCity = xmlItemDataXMLStream.ReadString();
                                break;

                            case "SHIPSTATE":
                                strShipState = xmlItemDataXMLStream.ReadString();
                                break;

                            case "SHIPZIP":
                                strShipZip = xmlItemDataXMLStream.ReadString();
                                break;

                            case "RGNCODE":
                                strRegionCode = xmlItemDataXMLStream.ReadString().ToUpper();
                                break;

                            case "ID":
                                strProductID = xmlItemDataXMLStream.ReadString();
                                break;

                            case "ADDITIONALNOTES":
                                strAdditionalNotes = xmlItemDataXMLStream.ReadString();
                                break;

                            case "SPLITOVERRIDE":
                                strSplitOverride = xmlItemDataXMLStream.ReadString().ToUpper();
                                break;

                            case "MAILCLASS":
                                strClassOfMail = xmlItemDataXMLStream.ReadString().ToUpper();
                                break;

                            case "SHIPPINGAMOUNT":
                                try
                                {
                                    dblShippingAmount = double.Parse(xmlItemDataXMLStream.ReadString());
                                }
                                catch
                                {
                                }
                                break;

                            case "TAXAMOUNT":
                                try
                                {
                                    dblTaxAmount = double.Parse(xmlItemDataXMLStream.ReadString());
                                }
                                catch
                                {
                                }
                                break;

                            default:
                                break;
                        }
                    }
                }

                int iFirstIndex = strItemDataXMLFullName.ToLower().LastIndexOf("order.") + 6;
                int iLastIndex = strItemDataXMLFullName.ToLower().LastIndexOf(".itemdata.");
                int iSubstringLength = iLastIndex - iFirstIndex;
                strStorefrontOrderID = strItemDataXMLFullName.Substring(iFirstIndex, iSubstringLength);

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////// CHECKING TO MAKE SURE ANY REQUIRED VALUES ARE POPULATED
                //////////////////////////////////////////////////////////////////////////////////////////
                if (string.IsNullOrEmpty(strOrderFirstName) || string.IsNullOrEmpty(strOrderLastName))
                {
                    LogFile("A First and Last name must be specified in the XML file.", true);
                    return false;
                }

                if (string.IsNullOrEmpty(strUserEmail) && string.IsNullOrEmpty(strDeliveryEmail))
                {
                    LogFile("A User Email or a Delivery Email must be specified in the XML file.", true);
                    return false;
                }

                if (string.IsNullOrEmpty(strOutputType) || ((!strOutputType.Equals("EMAILPDF")) && (!strOutputType.Equals("PRINTMAIL")) && (!strOutputType.Equals("PRINTSHIP")) && (!strOutputType.Equals("PICKANDPACK"))))
                {
                    LogFile("An unrecognized Output Type value of '" + strOutputType + "' was specified in the XML file.", true);
                    return false;
                }

                if ((!string.IsNullOrEmpty(strSplitOverride)) && (!strSplitOverride.Equals("HP")) && (!strSplitOverride.Equals("DIGI")))
                {
                    LogFile("An unrecognized Split Override value of '" + strSplitOverride + "' was specified in the XML file.", true);
                    return false;
                }

                if (strOutputType.Equals("PRINTMAIL") && !strClassOfMail.Equals("FIRST") && !strClassOfMail.Equals("STANDARD"))
                {
                    LogFile("An unrecognized Mail Class value of '" + strClassOfMail + "' was specified in the XML file.", true);
                    return false;
                }

                strProjectName = Path.GetFileNameWithoutExtension(strWFDNameOnly);

            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    Create Job Dates...    "

        static bool CreateJobDates()
        {
            try
            {
                if (!string.IsNullOrEmpty(strMailDate))
                {
                    if (DateTime.Parse(strMailDate) <= DateTime.Today)
                    {
                        switch (DateTime.Parse(strMailDate).AddDays(Convert.ToDouble(strDaysToAddForShipping)).DayOfWeek.ToString().ToUpper())
                        {
                            case "SATURDAY":
                                strDaysToAddForShipping = (Convert.ToInt16(strDaysToAddForShipping) + 2).ToString();
                                break;

                            case "SUNDAY":
                                strDaysToAddForShipping = (Convert.ToInt16(strDaysToAddForShipping) + 2).ToString();
                                break;

                            default:
                                break;
                        }

                        strMailDate = DateTime.Today.AddDays(Convert.ToDouble(strDaysToAddForShipping)).ToString("yyyy-MM-dd");
                    }
                }
                else
                {
                    switch (DateTime.Today.AddDays(Convert.ToDouble(strDaysToAddForShipping)).DayOfWeek.ToString().ToUpper())
                    {
                        case "SATURDAY":
                            strDaysToAddForShipping = (Convert.ToInt16(strDaysToAddForShipping) + 2).ToString();
                            break;

                        case "SUNDAY":
                            strDaysToAddForShipping = (Convert.ToInt16(strDaysToAddForShipping) + 2).ToString();
                            break;

                        default:
                            break;
                    }

                    strDeliveryDate = DateTime.Today.AddDays(Convert.ToDouble(strDaysToAddForShipping)).ToString("yyyy-MM-dd");
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    Parse Project Config_ORIGINAL...    "
        
        //////////////////////////////////////////
        /* NO LONGER USED
        //////////////////////////////////////////
        static bool ParseProjectConfig_ORIGINAL()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strConfigSwitch = "";
            string strProjectConfigLine;
            char[] aryProjectConfigLine;
            string strProjectConfigFile = strProjectConfigFolder + strProjectName + ".config";
            StreamReader streamProjectConfig;
            
            //////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////////////////////////////// PARSING THE PROJECT CONFIG FILE
            //////////////////////////////////////////////////////////////////////////////////////////   
            try
            {
                try
                {
                    streamProjectConfig = File.OpenText(strProjectConfigFile);
                }
                catch
                {
                    LogFile("Project Config file - " + strProjectConfigFile + " - cannot be found/opened.", true);
                    return false;
                }

                while (!streamProjectConfig.EndOfStream)
                {
                    strProjectConfigLine = streamProjectConfig.ReadLine();

                    if ((!(strProjectConfigLine.StartsWith("--"))) && (!(strProjectConfigLine.Trim().Equals(""))))
                    {
                        aryProjectConfigLine = strProjectConfigLine.ToCharArray();

                        for (int i = 0; i < aryProjectConfigLine.Length; i++)
                        {
                            if (aryProjectConfigLine[i].Equals(':'))
                            {
                                strConfigSwitch = strProjectConfigLine.Substring(0, i);
                                break;
                            }
                        }

                        switch (strConfigSwitch)
                        {
                            case "CASING":
                                strDataCasing = strProjectConfigLine.Replace("CASING:", "").Trim().ToUpper();
                                break;

                            case "NCOA-PAF":
                                strNCOACompany = strProjectConfigLine.Replace("NCOA-PAF:", "").Trim();
                                break;

                            case "MULTITRAC":
                                bolMultitrac = strProjectConfigLine.Replace("MULTITRAC:", "").Trim().ToUpper().Equals("TRUE") ? true : false;
                                break;

                            case "GENERATE-CLIENT-ID":
                                bolGenerateClientID = strProjectConfigLine.Replace("GENERATE-CLIENT-ID:", "").Trim().ToUpper().Equals("TRUE") ? true : false;
                                break;

                            case "IMPORT":
                                strImportName = strProjectConfigLine.Replace("IMPORT:", "").Trim();
                                break;

                            case "PRESORT":
                                strPresortSettings = strProjectConfigLine.Replace("PRESORT:", "").Trim();
                                break;

                            case "DEDUPE":
                                strDedupeType = strProjectConfigLine.Replace("DEDUPE:", "").Trim().ToUpper();
                                break;

                            case "OUTPUT":
                                strOutputLabel = strProjectConfigLine.Replace("OUTPUT:", "").Trim();
                                break;

                            case "MAIL-CLASS":
                                strClassOfMail = strProjectConfigLine.Replace("MAIL-CLASS:", "").Trim().ToUpper();
                                break;

                            case "MULTITRAC-CLIENT-ID":
                                strMultitracClientID = strProjectConfigLine.Replace("MULTITRAC-CLIENT-ID:", "").Trim();
                                break;

                            case "DATA-SPLIT-VALUE":
                                iDataSplitValue = int.Parse(strProjectConfigLine.Replace("DATA-SPLIT-VALUE:", "").Trim());
                                break;

                            case "APOGEE-CLIENT-FOLDER":
                                strApogeeClientFolder = strProjectConfigLine.Replace("APOGEE-CLIENT-FOLDER:", "").Trim();
                                break;

                            case "SEND-CLIENT-EMAIL":
                                bolSendClientEmail = strProjectConfigLine.Replace("SEND-CLIENT-EMAIL:", "").Trim().ToUpper().Equals("TRUE") ? true : false;
                                break;

                            default:
                                break;
                        }
                    }
                }

                streamProjectConfig.Close();
                streamProjectConfig.Dispose();
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }
        */
 
        #endregion


        #region "    Parse Project Config...    "

        static bool ParseProjectConfig()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strSQLCommand = "SELECT * FROM " + strSQLProjectConfigTable + " WHERE Project_Name = '" + strProjectName + "'";
            
            SqlConnection sqlDBConnection = new SqlConnection(strSQLConnection);
            SqlCommand sqlDBCommand = new SqlCommand(strSQLCommand, sqlDBConnection);
            SqlDataAdapter sqlDBDataAdapter = new SqlDataAdapter(sqlDBCommand);
            DataSet dataSQLTable = new DataSet();

            try
            {
                sqlDBDataAdapter.Fill(dataSQLTable);

                foreach (DataRow dataSQLTableRow in dataSQLTable.Tables[0].Rows)
                {
                    strCompanyName = dataSQLTableRow["Company_Name"].ToString().Trim();
                    strDataCasing = dataSQLTableRow["Data_Casing"].ToString().Trim().ToUpper();
                    strNCOACompany = dataSQLTableRow["NCOA_PAF"].ToString().Trim();
                    bolMultitrac = dataSQLTableRow["Multitrac_Job"].ToString().Trim().ToUpper().Equals("Y") ? true : false;
                    bolGenerateClientID = dataSQLTableRow["Generate_Client_ID"].ToString().Trim().ToUpper().Equals("Y") ? true : false;
                    strImportName = dataSQLTableRow["Import_Settings"].ToString().Trim();
                    strFirstClassPresortSettings = dataSQLTableRow["First_Class_Presort_Settings"].ToString().Trim();
                    strStandardPresortSettings = dataSQLTableRow["Standard_Presort_Settings"].ToString().Trim();
                    strDedupeType = dataSQLTableRow["Dedupe"].ToString().Trim().ToUpper();
                    strOutputLabel = dataSQLTableRow["Output_Label"].ToString().Trim();
                    strMultitracClientID = dataSQLTableRow["Multitrac_Client_ID"].ToString().Trim();
                    iDataSplitValue = int.Parse(dataSQLTableRow["Data_Split_Value"].ToString().Trim());
                    strApogeeClientFolder = dataSQLTableRow["Apogee_Client_Folder"].ToString().Trim();
                    bolSendClientEmail = dataSQLTableRow["Send_Client_Email"].ToString().Trim().ToUpper().Equals("Y") ? true : false;
                    strInternalEmailCC = dataSQLTableRow["Internal_Email_CC"].ToString().Trim();
                    strPrimaryKitKey = dataSQLTableRow["Primary_Kit_Key"].ToString().Trim().ToUpper();
                    strAdditionalKitKeyField = dataSQLTableRow["Additional_Kit_Key_Field"].ToString().Trim();
                    strMailingPrintOutput = dataSQLTableRow["Mailing_Print_Output"].ToString().Trim().ToUpper();
                    strSimplexOrDuplex = dataSQLTableRow["Simplex_or_Duplex"].ToString().Trim().ToUpper();
                    strDROrderPrefix = dataSQLTableRow["DR_Order_Prefix"].ToString().Trim().ToUpper();
                    strDRInventoryName = dataSQLTableRow["DR_Inventory_Name"].ToString().Trim();
                    dblOrderFee = double.Parse(dataSQLTableRow["Order_Fee"].ToString().Trim());
                    bolPaymentExpected = dataSQLTableRow["Payment_Expected"].ToString().Trim().ToUpper().Equals("Y") ? true : false;
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////// VERIFYING ALL FIELDS THAT HAVE SPECIFIC ACCEPTABLE VALUES
                //////////////////////////////////////////////////////////////////////////////////////////
                if (string.IsNullOrEmpty(strCompanyName))
                {
                    LogFile("A '" + strProjectName + "' project could not be found in the Project Config table.", true);
                    return false;
                }
                
                if (!string.IsNullOrEmpty(strDataCasing) && !strDataCasing.Equals("MIXED") && !strDataCasing.Equals("UPPER"))
                {
                    LogFile("Acceptable values for Data Casing are as follows:  MIXED or UPPER.", true);
                    return false;
                }

                if (!string.IsNullOrEmpty(strDedupeType) && !strDedupeType.Equals("NONE") && !strDedupeType.Equals("PLAYER-ID") && !strDedupeType.Equals("CARD-ID") && !strDedupeType.Equals("ACCOUNT-ID") && !strDedupeType.Equals("ONE-PER-PERSON") && !strDedupeType.Equals("ONE-PER-ADDRESS"))
                {
                    LogFile("Acceptable values for Dedupe are as follows:  NONE, PLAYER-ID, CARD-ID, ACCOUNT-ID, ONE-PER-PERSON, or ONE-PER-ADDRESS", true);
                    return false;
                }

                if (!string.IsNullOrEmpty(strMailingPrintOutput) && !strMailingPrintOutput.Equals("DIGITAL") && !strMailingPrintOutput.Equals("INKJET"))
                {
                    LogFile("Acceptable values for Mailing Print Output are as follows:  DIGITAL or INKJET", true);
                    return false;
                }

                if (!string.IsNullOrEmpty(strSimplexOrDuplex) && !strSimplexOrDuplex.Equals("SIMPLEX") && !strSimplexOrDuplex.Equals("DUPLEX"))
                {
                    LogFile("Acceptable values for Simplex Or Duplex are as follows:  SIMPLEX or DUPLEX.", true);
                    return false;
                }

                if (bolMultitrac && !strOutputLabel.ToUpper().Contains("MULTITRAC"))
                {
                    LogFile("Product '" + strProjectName + "' has been flagged as a MultiTrac job, but is not using a MultiTrac output label.", true);
                    return false;
                }

                if (!strApogeeClientFolder.Substring((strApogeeClientFolder.Length - 1),1).Equals(@"\"))
                {
                    strApogeeClientFolder = strApogeeClientFolder + @"\";
                }
                
                if (strClassOfMail.Equals("FIRST"))
                {
                    strPresortSettings = strFirstClassPresortSettings;
                }
                else
                {
                    strPresortSettings = strStandardPresortSettings;
                }

                if (strOutputType.Equals("PRINTMAIL"))
                {
                    if (string.IsNullOrEmpty(strPresortSettings))
                    {
                        LogFile("A Presort Settings value for the specified Mail Class - " + strClassOfMail + " - could not be found in the Project Config table.", true);
                        return false;
                    }
                }

                strWFDFullName = strStoreFrontFolder + strCompanyName + "\\WFDs\\Live\\" + strWFDNameOnly;
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


        #region "    Pace Process...    "

        //////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////// NONE OF THIS IS READY YET - PACE WILL BE INTEGRATED AT A LATER DATE
        //////////////////////////////////////////////////////////////////////////////////////////
        
        static bool PaceProcess()
        {
            try
            {
                if (bolRunInTestMode)
                {
                    strJobNumber = strDirectResponseOrderNumber;                    // DEBUGGING.
                    strClientCode = "STOREFRONT";                                   // DEBUGGING.
                    
                    //strJobNumber = strTestJobNumber;
                    //strClientCode = strTestClientCode;
                }
                else
                {
                    strJobNumber = strDirectResponseOrderNumber;
                    strClientCode = "STOREFRONT";
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    File Placement...    "

        static bool FilePlacement()
        {
            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////// CREATING THE JOB FOLDER AND ALL NEEDED SUBFOLDERS FOR PROCESSING
                //////////////////////////////////////////////////////////////////////////////////////////
                string strDataRootJobFolder = strDataJobsFolder + strJobNumber + "\\";
                string strGeneralRootJobFolder = strGeneralJobsFolder + strJobNumber + "\\";

                // Verifying that the job folder does not already exist.
                if (Directory.Exists(strDataRootJobFolder))
                {
                    LogFile("A job folder for " + strJobNumber + " already exists on 'data' share.", true);
                    return false;
                }

                if (Directory.Exists(strGeneralRootJobFolder))
                {
                    LogFile("A job folder for " + strJobNumber + " already exists on 'general' share.", true);
                    return false;
                }

                // Creating 'data' share job folder structure.
                strBCCJobFolder = strDataRootJobFolder + "BCC\\";
                strInputJobFolder = strDataRootJobFolder + "Input\\";
                strPresortDataJobFolder = strDataRootJobFolder + "PresortData\\";
                string strProgramsJobFolder = strDataRootJobFolder + "Programs\\";
                strWorkingJobFolder = strDataRootJobFolder + "Working\\";
                string strUnidentifiedFolder = strWorkingJobFolder + "Unidentified_Files\\";

                // Creating 'general' share job folder structure.
                strCSRJobFolder = strGeneralRootJobFolder + "CSR\\";
                strDocsJobFolder = strGeneralRootJobFolder + "Docs\\";
                strPNetImagesJobFolder = strGeneralRootJobFolder + "PNetImages\\";
                string strPrepressJobFolder = strGeneralRootJobFolder + "Prepress\\";
                strProductionJobFolder = strGeneralRootJobFolder + "Production\\";

                // Creating the specified folder/subfolders.
                Directory.CreateDirectory(strBCCJobFolder);
                Directory.CreateDirectory(strInputJobFolder);
                Directory.CreateDirectory(strPresortDataJobFolder);
                Directory.CreateDirectory(strProgramsJobFolder);
                Directory.CreateDirectory(strWorkingJobFolder);
                Directory.CreateDirectory(strUnidentifiedFolder);
                
                Directory.CreateDirectory(strCSRJobFolder);
                Directory.CreateDirectory(strDocsJobFolder);
                Directory.CreateDirectory(strPNetImagesJobFolder);
                Directory.CreateDirectory(strPrepressJobFolder);
                Directory.CreateDirectory(strProductionJobFolder);

                // Moving the input data files from the decompressed folder to their respective job folder.
                File.Move(strInputXMLName, strInputJobFolder + Path.GetFileName(strInputXMLName));
                File.Move(strItemDataXMLFullName, strInputJobFolder + strItemDataXMLNameOnly);

                if (!strInputDataName.Equals(""))
                {
                    File.Move(strInputDataName, strInputJobFolder + Path.GetFileName(strInputDataName));
                }

                if (lstImageFiles.Count > 0)
                {
                    foreach (string strImageName in lstImageFiles)
                    {
                        File.Move(strDecompressedZipFolder + strImageName, strPNetImagesJobFolder + strImageName);
                    }
                }

                // Moving any unidentified files into the 'unidentified' folder. If all files are identified, folder will be removed.
                ProcessStartInfo cmdMoveFilesStartInfo = new ProcessStartInfo();
                cmdMoveFilesStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cmdMoveFilesStartInfo.RedirectStandardError = true;
                cmdMoveFilesStartInfo.RedirectStandardOutput = true;
                cmdMoveFilesStartInfo.UseShellExecute = false;
                cmdMoveFilesStartInfo.CreateNoWindow = true;
                cmdMoveFilesStartInfo.FileName = "cmd.exe";
                cmdMoveFilesStartInfo.Arguments = "/C move /Y \"" + strDecompressedZipFolder + "*.*\" \"" + strUnidentifiedFolder + "\"";

                Process cmdMoveFiles = new Process();
                cmdMoveFiles = Process.Start(cmdMoveFilesStartInfo);
                cmdMoveFiles.WaitForExit();
                cmdMoveFiles.Dispose();

                if (Directory.GetFiles(strUnidentifiedFolder).Length == 0)
                {
                    Directory.Delete(strUnidentifiedFolder);
                }

                // Resetting file names to their new locations.
                strInputDataName = strInputJobFolder + Path.GetFileName(strInputDataName);
                strItemDataXMLFullName = strInputJobFolder + Path.GetFileName(strItemDataXMLFullName);
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    Verify Zip and Region...    "

        static bool VerifyZipAndRegion()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////                
            string[] aryTableParsedLine;
            string[] aryDataParsedLine;
            string strDataLine;
            string strTableLine;
            string strEditedDataFile = strInputDataName.ToUpper().Replace(".CSV", "_EDITED.CSV");
            string strCurrentZip;
            string strDataZipAndRegion;
            string strTableZipAndRegion;
            int iZipFieldIndex = 0;
            int iDataFields;
            bool bolFoundRegionMatch;
            int iInputRecords;
            int iEditedRecords;
            int iMismatchRecords;
            bool bolZipFieldFound = false;
            strRegionErrorsRemoved = strWorkingJobFolder + strJobNumber + " - Region Mismatches Removed.csv";

            StreamReader streamInitialFileScan = new StreamReader(strInputDataName);
            StreamReader streamTableFile = new StreamReader(strZipToRegionTable);
            StreamWriter streamEditedDataFile = new StreamWriter(strEditedDataFile);
            StreamWriter streamRegionMismatches = new StreamWriter(strRegionErrorsRemoved);

            TextFieldParser parseDataFile = new TextFieldParser(strInputDataName);
            parseDataFile.TextFieldType = FieldType.Delimited;
            parseDataFile.SetDelimiters(",");

            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////// DETERMINING WHICH FIELD IN THE INPUT DATA CONTAINS THE ZIP CODE
                ////////////////////////////////////////////////////////////////////////////////////////// 
                strDataLine = streamInitialFileScan.ReadLine();  
                aryDataParsedLine = strDataLine.Split(',');
                iDataFields = aryDataParsedLine.Length;

                for (int j = 0; j < iDataFields; j++)
                {
                    if (aryDataParsedLine[j].ToString().ToUpper().Contains("ZIP"))
                    {
                        bolZipFieldFound = true;
                        streamEditedDataFile.WriteLine(strDataLine);
                        iZipFieldIndex = j;
                            
                        break;
                    }
                }

                streamInitialFileScan.Close();
                streamInitialFileScan.Dispose();

                // Verifying that a zip code field exists in the input data file.
                if (!bolZipFieldFound)
                {
                    LogFile("A Zip field is not included in the input data file.", true);
                    return false;
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////// TESTING EACH RECORD AGAINST THE ZIP-REGION TABLE
                ////////////////////////////////////////////////////////////////////////////////////////// 
                while (!parseDataFile.EndOfData)
                {
                    bolFoundRegionMatch = false;

                    strDataLine = parseDataFile.PeekChars(Int32.MaxValue);
                    aryDataParsedLine = parseDataFile.ReadFields();
                    
                    // Capturing the zip and region combination from the current record.
                    strCurrentZip = aryDataParsedLine[iZipFieldIndex].ToString().Trim();

                    if (strCurrentZip.Length > 5)
                    {
                        strCurrentZip = strCurrentZip.Substring(0,5);
                    }

                    strDataZipAndRegion = strCurrentZip + strRegionCode;

                    // Looping through the Zip and Region Lookup Table to see if a zip is within a valid region.
                    while (!streamTableFile.EndOfStream)
                    {
                        strTableLine = streamTableFile.ReadLine();
                        aryTableParsedLine = strTableLine.Split(',');

                        strTableZipAndRegion = aryTableParsedLine[0].ToString() + aryTableParsedLine[2].ToString().ToUpper().Trim();

                        if (strDataZipAndRegion == strTableZipAndRegion)
                        {
                            bolFoundRegionMatch = true;

                            break;
                        }
                    }

                    if (bolFoundRegionMatch)
                    {
                        streamEditedDataFile.WriteLine(strDataLine);
                    }
                    else
                    {
                        streamRegionMismatches.WriteLine(strDataLine);
                    }

                    streamTableFile.DiscardBufferedData();
                    streamTableFile.BaseStream.Position = 0;
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }
            finally
            {
                streamEditedDataFile.Close();
                streamEditedDataFile.Dispose();
                streamTableFile.Close();
                streamTableFile.Dispose();
                streamRegionMismatches.Close();
                streamRegionMismatches.Dispose();
                parseDataFile.Close();
                parseDataFile.Dispose();
            }

            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////// DETERMINING IF THE HIGHMARK THRESHOLD HAS BEEN REACHED
            ////////////////////////////////////////////////////////////////////////////////////////// 
            try
            {
                // Calculating total number of input records.
                iInputRecords = File.ReadAllLines(strInputDataName).Length - 1;

                // Calculating total number of edited records.
                iEditedRecords = File.ReadAllLines(strEditedDataFile).Length - 1;

                // Calculating total number of mismatch records.
                iMismatchRecords = File.ReadAllLines(strRegionErrorsRemoved).Length - 1;

                if ((((decimal)iEditedRecords / (decimal)iInputRecords) * 100) <= iZipToRegionMismatchThreshold)
                {
                    bolRegionMismatchThresholdReached = true;

                    SendEmail("HIGHMARK");
                    LogFile("At least " + (100 - iZipToRegionMismatchThreshold).ToString() + "% of records submitted for processing were removed as Region-Zip mismatches.", true);
                    return false;
                }
                else
                {
                    if (iMismatchRecords > 1)
                    {
                        SendEmail("HIGHMARK");
                    }
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            strInputDataName = strEditedDataFile;
            return true;
        }

        #endregion


        #region "    BCC Process...    "

        static bool BCCProcess()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                string strMJBProcess = strTaskmasterJobsFolder + strProjectName + ".mjb";
                string strCommandLine = "-j \"" + Path.GetFileName(strMJBProcess) + "\" -u \"AUTO\" -w \"AUTO\" -r";

                // MJB and MDS Values
                string strMailDatSettings = strMailDatSettingsFolder + strProjectName + ".mds";
                strBCCListName = strJobNumber + " - " + strProjectName;
                strBCCDatabase = strBCCJobFolder + strBCCListName + ".dbf";
                strMailDate = DateTime.Parse(strMailDate).ToString("MM/dd/yyyy");
                
                // Postal Paperwork
                string strPostageStatementPDF = strDocsJobFolder + strJobNumber + " - Postage Statement.pdf";
                string strQualReportPDF = strDocsJobFolder + strJobNumber + " - Qualification Report.pdf";
                string strPalletPlacardsPDF = strDocsJobFolder + strJobNumber + " - Pallet Placards.pdf";
                string strTrayTagsPDF = strDocsJobFolder + strJobNumber + " - Tray Tags.pdf";

                // Output values
                strInvalidsRemoved = strWorkingJobFolder + strJobNumber + " - Invalids Removed.xls";
                strDupesRemoved = strWorkingJobFolder + strJobNumber + " - Dupes Removed.xls";
                strMultitracInput = strWorkingJobFolder + strJobNumber + " - Multitrac Input.csv";
                strMultitracOutput = strWorkingJobFolder + strJobNumber + " - Multitrac Output.csv";
                strSortedData = strPresortDataJobFolder + strJobNumber + " - Sorted.csv";

                bool bolMultitracSuccess = false;
                Process cmdFirstPassProcess = new Process();
                Process cmdSecondPassProcess = new Process();
                Process cmdPaperworkProcess = new Process();
                ProcessStartInfo cmdFirstPassStartInfo = new ProcessStartInfo();
                ProcessStartInfo cmdSecondPassStartInfo = new ProcessStartInfo();
                ProcessStartInfo cmdPaperworkStartInfo = new ProcessStartInfo();
                StreamWriter streamMailDatSettings;
                StreamWriter streamFirstPassMJBProcess;
                StreamWriter streamSecondPassMJBProcess;
                //StreamWriter streamPaperworkMJBProcess;
                string[] aryInputRecords;
                int iInputRecords;
                string[] aryMailingRecords;                

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////// CALCULATING TOTAL NUMBER OF INPUT RECORDS
                //////////////////////////////////////////////////////////////////////////////////////////
                aryInputRecords = File.ReadAllLines(strInputDataName);
                iInputRecords = aryInputRecords.Length - 1;

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////////// CREATING FIRST PASS MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                streamFirstPassMJBProcess = new StreamWriter(strMJBProcess, false);

                streamFirstPassMJBProcess.WriteLine("[NEWLISTTEMPLATE-1]");
                streamFirstPassMJBProcess.WriteLine("DESCRIPTION=\"" + strBCCListName + "\"");
                streamFirstPassMJBProcess.WriteLine("FILENAME=\"" + strBCCDatabase + "\"");
                streamFirstPassMJBProcess.WriteLine("OVERWRITE=Y");
                streamFirstPassMJBProcess.WriteLine("GROUP=\"AUTO\"");
                streamFirstPassMJBProcess.WriteLine("SETTINGS=\"Heeter Standard Template\"");
                streamFirstPassMJBProcess.WriteLine("USEINDEXES=Y");
                streamFirstPassMJBProcess.WriteLine("");

                streamFirstPassMJBProcess.WriteLine("[IMPORT-2]");
                streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamFirstPassMJBProcess.WriteLine("SETTINGS=\"" + strImportName + "\"");
                streamFirstPassMJBProcess.WriteLine("FILENAME=\"" + strInputDataName + "\"");
                streamFirstPassMJBProcess.WriteLine("");

                /*
                streamFirstPassMJBProcess.WriteLine("[ENCODE-3]");
                streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamFirstPassMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamFirstPassMJBProcess.WriteLine("ADDRESSGROUPS=\"MAIN\"");
                streamFirstPassMJBProcess.WriteLine("SWAP=Y");
                streamFirstPassMJBProcess.WriteLine("STANDARDIZEADDRESS=Y");
                streamFirstPassMJBProcess.WriteLine("STANDARDIZECITY=Y");
                streamFirstPassMJBProcess.WriteLine("ABBREVIATECITY=N");
                streamFirstPassMJBProcess.WriteLine("IGNORENONUSPS=Y");
                streamFirstPassMJBProcess.WriteLine("EXTENDEDMATCHING=N");
                streamFirstPassMJBProcess.WriteLine("CASE=\"ASIS\"");
                streamFirstPassMJBProcess.WriteLine("FIRMASIS=Y");
                streamFirstPassMJBProcess.WriteLine("ZIP5CHECKDIGIT=N");
                streamFirstPassMJBProcess.WriteLine("SUMMARYPAGE=N");
                streamFirstPassMJBProcess.WriteLine("NDIREPORT=N");
                streamFirstPassMJBProcess.WriteLine("COPIES=0");
                streamFirstPassMJBProcess.WriteLine("PREPAREDFOR=NONE");
                streamFirstPassMJBProcess.WriteLine("");
                */

                streamFirstPassMJBProcess.WriteLine("[MODIFY-4]");
                streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamFirstPassMJBProcess.WriteLine("SETTINGS=\"Standard Address Block Filter\"");
                streamFirstPassMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamFirstPassMJBProcess.WriteLine("");

                if (iInputRecords >= 100)
                {
                    streamFirstPassMJBProcess.WriteLine("[DATASERVICES-5]");
                    streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamFirstPassMJBProcess.WriteLine("PROCESSES=\"FSP\"");
                    streamFirstPassMJBProcess.WriteLine("CLASSOFMAIL=\"F\"");
                    streamFirstPassMJBProcess.WriteLine("MAILINGZIPCODE=\"15290\"");
                    streamFirstPassMJBProcess.WriteLine("LISTOWNER=\"" + strNCOACompany + "\"");
                    streamFirstPassMJBProcess.WriteLine("PREPAID=Y");
                    streamFirstPassMJBProcess.WriteLine("EXTENDEDMATCHING=Y");
                    streamFirstPassMJBProcess.WriteLine("PAFELECTRONIC=N");
                    streamFirstPassMJBProcess.WriteLine("JOBPASSWORD=120000001EE18648DBD198CC853206D2E93C8604120DFBA79417C9E2A68A8C523881F5D30316FA979FC95A5484F12977A9569A570F338049AB987585C4609D62677D8B901FC0672C8E037EC8212AEEC9D3BB8E0E");
                    streamFirstPassMJBProcess.WriteLine("ADDRESSGROUPS=\"MAIN\"");
                    streamFirstPassMJBProcess.WriteLine("ORDERTERMSACCEPTED=Y");
                    streamFirstPassMJBProcess.WriteLine("CASE=\"AUTO\"");
                    streamFirstPassMJBProcess.WriteLine("COAAUDITEXPORT=Y");
                    streamFirstPassMJBProcess.WriteLine("STANDARDIZECITY=N");
                    streamFirstPassMJBProcess.WriteLine("HIDERETURNCODES=\"10 11 12 13 17 21 26 27 28 33 98\"");
                    streamFirstPassMJBProcess.WriteLine("");
                }

                streamFirstPassMJBProcess.WriteLine("[EXPORT-6]");
                streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamFirstPassMJBProcess.WriteLine("FILENAME=\"" + strInvalidsRemoved + "\"");
                streamFirstPassMJBProcess.WriteLine("SETTINGS=\"Dupes or Invalids Export\"");
                streamFirstPassMJBProcess.WriteLine("SELECTIVITY=\"Hidden Record\"");
                streamFirstPassMJBProcess.WriteLine("INDEX=NONE");
                streamFirstPassMJBProcess.WriteLine("OVERWRITE=Y");
                streamFirstPassMJBProcess.WriteLine("");

                streamFirstPassMJBProcess.WriteLine("[DELETEHIDDENRECORDS-7]");
                streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamFirstPassMJBProcess.WriteLine("");

                streamFirstPassMJBProcess.WriteLine("[MODIFY-8]");
                streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamFirstPassMJBProcess.WriteLine("SETTINGS=\"" + ((strDataCasing.Equals("UPPER")) ? "CASING: Upper Case Address Block" : "CASING: Mixed Case Address Block") + "\"");
                streamFirstPassMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamFirstPassMJBProcess.WriteLine("");

                if (!strDedupeType.Equals("NONE"))
                {
                    string strDedupeName = "";

                    switch (strDedupeType)
                    {
                        case "PLAYER-ID":
                            strDedupeName = "Dedupe: Player ID";
                            break;

                        case "CARD-ID":
                            strDedupeName = "Dedupe: Card ID";
                            break;

                        case "ACCOUNT-ID":
                            strDedupeName = "Dedupe: Account ID";
                            break;

                        case "ONE-PER-PERSON":
                            strDedupeName = "Dedupe: Fullname, Address1, and Zip";
                            break;

                        case "ONE-PER-ADDRESS":
                            strDedupeName = "Dedupe: Address1 and Zip";
                            break;

                        default:
                            break;
                    }

                    streamFirstPassMJBProcess.WriteLine("[DEDUPE-9]");
                    streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamFirstPassMJBProcess.WriteLine("SETTINGS=\"" + strDedupeName + "\"");
                    streamFirstPassMJBProcess.WriteLine("");

                    streamFirstPassMJBProcess.WriteLine("[EXPORT-10]");
                    streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamFirstPassMJBProcess.WriteLine("FILENAME=\"" + strDupesRemoved + "\"");
                    streamFirstPassMJBProcess.WriteLine("SETTINGS=\"Dupes or Invalids Export\"");
                    streamFirstPassMJBProcess.WriteLine("SELECTIVITY=\"Hidden Record\"");
                    streamFirstPassMJBProcess.WriteLine("INDEX=NONE");
                    streamFirstPassMJBProcess.WriteLine("OVERWRITE=Y");
                    streamFirstPassMJBProcess.WriteLine("");

                    streamFirstPassMJBProcess.WriteLine("[DELETEHIDDENRECORDS-11]");
                    streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamFirstPassMJBProcess.WriteLine("");
                }

                if (bolMultitrac)
                {
                    if (bolGenerateClientID)
                    {
                        streamFirstPassMJBProcess.WriteLine("[MODIFY-12]");
                        streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                        streamFirstPassMJBProcess.WriteLine("SETTINGS=\"MULTITRAC: Generate Client ID\"");
                        streamFirstPassMJBProcess.WriteLine("SELECTIVITY=NONE");
                        streamFirstPassMJBProcess.WriteLine("");
                    }

                    streamFirstPassMJBProcess.WriteLine("[MODIFY-13]");
                    streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamFirstPassMJBProcess.WriteLine("SETTINGS=\"MULTITRAC: Fix ZIP+4-DP Combo For Import\"");
                    streamFirstPassMJBProcess.WriteLine("SELECTIVITY=NONE");
                    streamFirstPassMJBProcess.WriteLine("");

                    streamFirstPassMJBProcess.WriteLine("[EXPORT-14]");
                    streamFirstPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamFirstPassMJBProcess.WriteLine("FILENAME=\"" + strMultitracInput + "\"");
                    streamFirstPassMJBProcess.WriteLine("SETTINGS=\"MULTITRAC: MT Input File\"");
                    streamFirstPassMJBProcess.WriteLine("SELECTIVITY=NONE");
                    streamFirstPassMJBProcess.WriteLine("INDEX=NONE");
                    streamFirstPassMJBProcess.WriteLine("OVERWRITE=Y");
                    streamFirstPassMJBProcess.WriteLine("");
                }

                streamFirstPassMJBProcess.Close();
                streamFirstPassMJBProcess.Dispose();

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////// RUNNING FIRST PASS MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdFirstPassStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cmdFirstPassStartInfo.RedirectStandardError = false;
                cmdFirstPassStartInfo.RedirectStandardOutput = false;
                cmdFirstPassStartInfo.UseShellExecute = true;
                cmdFirstPassStartInfo.CreateNoWindow = true;
                cmdFirstPassStartInfo.FileName = strBCCMailManEXE;
                cmdFirstPassStartInfo.Arguments = strCommandLine;

                cmdFirstPassProcess = Process.Start(cmdFirstPassStartInfo);
                cmdFirstPassProcess.WaitForExit();
                cmdFirstPassProcess.Dispose();

                File.Delete(strMJBProcess);

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////// PROCESSING DATA THROUGH MULTITRAC, IF NEEDED FOR CLIENT
                //////////////////////////////////////////////////////////////////////////////////////////
                if (bolMultitrac)
                {
                    bolMultitracSuccess = MultitracProcess("PROCESS_JOB");
                    if (!bolMultitracSuccess) return false;
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////// CREATING MAIL.DAT SETTINGS FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                streamMailDatSettings = new StreamWriter(strMailDatSettings, false);

                streamMailDatSettings.WriteLine("  FileSetName = '" + strJobNumber + "'");
                streamMailDatSettings.WriteLine("  FileSource = 'Heeter Direct'");
                streamMailDatSettings.WriteLine("  Location = '" + strDocsJobFolder + "'");
                streamMailDatSettings.WriteLine("  JobNumber = '" + strJobNumber + "'");
                streamMailDatSettings.WriteLine("  JobName = '" + strJobNumber + "'");
                streamMailDatSettings.WriteLine("  ContactName = '" + strMailDatContactName + "'");
                streamMailDatSettings.WriteLine("  ContactEmail = '" + strMailDatContactEmail + "'");
                streamMailDatSettings.WriteLine("  ContactPhone = '" + strMailDatContactPhone + "'");
                streamMailDatSettings.WriteLine("  SegmentingCriteria = '" + strJobNumber + "'");
                streamMailDatSettings.WriteLine("  PostalOne = True");
                streamMailDatSettings.WriteLine("  IMRTable = False");
                streamMailDatSettings.WriteLine("  WSRTable = False");
                streamMailDatSettings.WriteLine("  PDRTable = True");
                streamMailDatSettings.WriteLine("  SFRTable = False");
                streamMailDatSettings.WriteLine("  SNRTable = False");
                streamMailDatSettings.WriteLine("  MSRTable = False");
                streamMailDatSettings.WriteLine("  MIRTable = False");
                streamMailDatSettings.WriteLine("  PBCTable = False");
                streamMailDatSettings.WriteLine("  FixedBatch = False");
                streamMailDatSettings.WriteLine("  BatchSize = 300");
                streamMailDatSettings.WriteLine("  UseMailDatFolder = True");
                streamMailDatSettings.WriteLine("  UseJobNumber = True");
                streamMailDatSettings.WriteLine("  ZipDatabase = True");
                streamMailDatSettings.WriteLine("  DeleteDatabaseAfterZipping = False");
                streamMailDatSettings.WriteLine("  SegmentDescription = '" + strJobNumber + "'");
                streamMailDatSettings.WriteLine("  MailingFacility = 'Pittsburgh PA'");
                streamMailDatSettings.WriteLine("  MailingFacilityZIP4 = '152901001'");
                streamMailDatSettings.WriteLine("  NumericDisplayContainerID = False");
                streamMailDatSettings.WriteLine("  DetachedAddressLabels = False");
                streamMailDatSettings.WriteLine("  UseConfirmBarcode = False");
                streamMailDatSettings.WriteLine("  Services = False");
                streamMailDatSettings.WriteLine("  ContainerInfo = True");
                streamMailDatSettings.WriteLine("  EDocSenderCrid = '4057074'");
                streamMailDatSettings.WriteLine("  BypassSeamlessAcceptance = False");
                streamMailDatSettings.WriteLine("  USEIMpb = False");
                streamMailDatSettings.WriteLine("  MailPieceName = '" + (strClassOfMail.Equals("FIRST") ? "FC" : "STD") + "'");
                streamMailDatSettings.WriteLine("  Enclosure = False");
                streamMailDatSettings.WriteLine("  RideAlong = False");
                streamMailDatSettings.WriteLine("  RepositionableNote = False");
                streamMailDatSettings.WriteLine("  ComponentDescription = '" + strJobNumber + "'");
                streamMailDatSettings.WriteLine("  ContentOfMail = '  '");
                streamMailDatSettings.WriteLine("  PostalPriceIncentiveType = '  '");
                streamMailDatSettings.WriteLine("  EnclosureBulkInsurance = False");
                streamMailDatSettings.WriteLine("  ActualEntryDiffers = False");
                streamMailDatSettings.WriteLine("  InHomeRange = 0");
                streamMailDatSettings.WriteLine("  EInduction = False");
                streamMailDatSettings.WriteLine("  AcceptMisshipped = False");
                streamMailDatSettings.WriteLine("  ContainerTags = True");
                streamMailDatSettings.WriteLine("  MailerMailerLocation = 'Pittsburgh PA'");
                streamMailDatSettings.WriteLine("  InformationLine = True");
                //streamMailDatSettings.WriteLine("  UserInformationLine = '%T  %N  %U  %K  %P'");
                //streamMailDatSettings.WriteLine("  UserInformationLine2 = '%T  %X  %P  %N'");
                streamMailDatSettings.WriteLine("  ResetUserInformationSackNumber = False");
                streamMailDatSettings.WriteLine("  IMContainerTags = True");
                streamMailDatSettings.WriteLine("  IMContainerTagsMailerID = '" + strHeeterMailerID + "'");
                streamMailDatSettings.WriteLine("  Newspaper = False");
                streamMailDatSettings.WriteLine("  PostagePaymentMethod = 'P'");
                streamMailDatSettings.WriteLine("  ResetPackageID = False");
                streamMailDatSettings.WriteLine("  SequentialPieceID = False");
                streamMailDatSettings.WriteLine("  MailPieceStatus = False");
                streamMailDatSettings.WriteLine("  AddACSKeylineCheckDigit = False");
                streamMailDatSettings.WriteLine("  BulkInsurance = False");
                streamMailDatSettings.WriteLine("  Confirm14DigitPlanetBarcode = False");
                streamMailDatSettings.WriteLine("  IMBUseServicesExpression = False");

                if (bolMultitrac)
                {
                    streamMailDatSettings.WriteLine("  IMBMailerIDExpression = '[Mailer ID]'");
                    streamMailDatSettings.WriteLine("  IMBSerialNumberExpression = '[IMB Sequence]'");
                }
                else
                {
                    streamMailDatSettings.WriteLine("  IMBMailerIDExpression = '" + strHeeterMailerID + "'");
                    streamMailDatSettings.WriteLine("  IMBSerialNumberExpression = ''");
                }

                streamMailDatSettings.WriteLine("  IMBConfirmSampling = False");
                streamMailDatSettings.WriteLine("  IMpbUseServicesExpression = False");
                streamMailDatSettings.WriteLine("  IMpbZIPFormat = '9'");
                streamMailDatSettings.WriteLine("  IMpbNonUSPSValid = False");
                streamMailDatSettings.WriteLine("  InputMode = 'OVERWRITE'");
                streamMailDatSettings.WriteLine("  FileSetStatus = 'O'");
                streamMailDatSettings.WriteLine("  Version = '" + strMailDatVersion + "'");
                streamMailDatSettings.WriteLine("  MoveUpdateDate = ''");
                streamMailDatSettings.WriteLine("  WalkSequenceDate = ''");
                streamMailDatSettings.WriteLine("  SubstitutedContainerPrep = ' '");
                streamMailDatSettings.WriteLine("  MailPieceAdStatus = 'N'");
                streamMailDatSettings.WriteLine("  ComponentRateType = 'R'");
                streamMailDatSettings.WriteLine("  RideAlongProcessingCategory = 'LT'");
                streamMailDatSettings.WriteLine("  RideAlongWeightStatus = 'F'");
                streamMailDatSettings.WriteLine("  RideAlongWeightSource = 'C'");
                streamMailDatSettings.WriteLine("  EnclosureType = 'N'");
                streamMailDatSettings.WriteLine("  EnclosureRateType = 'X'");
                streamMailDatSettings.WriteLine("  EnclosureProcessingCategory = 'LT'");
                streamMailDatSettings.WriteLine("  EnclosureWeightStatus = 'F'");
                streamMailDatSettings.WriteLine("  EnclosureWeightSource = 'C'");
                streamMailDatSettings.WriteLine("  ActualEntryFacilityType = ' '");
                streamMailDatSettings.WriteLine("  ContainerStatus = 'R'");
                streamMailDatSettings.WriteLine("  ShipScheduledDateTime = '" + strMailDate + "'");
                streamMailDatSettings.WriteLine("  ShipDate = '" + strMailDate + "'");
                streamMailDatSettings.WriteLine("  InHomeDate = ''");
                streamMailDatSettings.WriteLine("  StatementDateTime = '" + strMailDate + " 00:00'");
                streamMailDatSettings.WriteLine("  InductionDateTime = ''");
                streamMailDatSettings.WriteLine("  InductionActualDateTime = ''");
                streamMailDatSettings.WriteLine("  InternalDate = ''");
                streamMailDatSettings.WriteLine("  PendingPeriodical = 'N'");
                streamMailDatSettings.WriteLine("  ContainerChargeMethod = '2'");
                streamMailDatSettings.WriteLine("  IssueDate = ''");
                streamMailDatSettings.WriteLine("  ComponentAdStatus = 'N'");
                streamMailDatSettings.WriteLine("  ServiceSet = ''");
                streamMailDatSettings.WriteLine("  ConfirmService = '22'");
                streamMailDatSettings.WriteLine("  PlanetBarcodeOption = 'S'");
                streamMailDatSettings.WriteLine("  SeedType = 'R'");
                streamMailDatSettings.WriteLine("  ZIP4EncodingDate = ''");

                if (bolMultitrac)
                {
                    streamMailDatSettings.WriteLine("  IMBServices = 'DCM FLS'");
                }
                else
                {
                    streamMailDatSettings.WriteLine("  IMBServices = 'FLS'");
                }

                streamMailDatSettings.WriteLine("  PickupScheduledDateTime = ''");
                streamMailDatSettings.WriteLine("  PickupDateTime = ''");
                streamMailDatSettings.WriteLine("  MoveUpdateMethod = '0'");
                streamMailDatSettings.WriteLine("  USPSPickupMailing = 'N'");
                streamMailDatSettings.WriteLine("  USPSPickup = 'N'");
                streamMailDatSettings.WriteLine("  SeamlessAcceptanceIndicator = ' '");
                streamMailDatSettings.WriteLine("  FullServiceParticipation = 'F'");
                streamMailDatSettings.WriteLine("  SASPPreparationOption = ' '");
                streamMailDatSettings.WriteLine("  ACSKeyline = 'N'");
                streamMailDatSettings.WriteLine("  DetachedMailingLabel = ' '");
                streamMailDatSettings.WriteLine("  CharacteristicFee = '  '");
                streamMailDatSettings.WriteLine("  PostagePaymentOption = 'D'");

                streamMailDatSettings.Close();
                streamMailDatSettings.Dispose();

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////// CREATING SECOND PASS MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                streamSecondPassMJBProcess = new StreamWriter(strMJBProcess, false);

                if (bolMultitrac)
                {
                    streamSecondPassMJBProcess.WriteLine("[IMPORT-1]");
                    streamSecondPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                    streamSecondPassMJBProcess.WriteLine("SETTINGS=\"MULTITRAC: Output Data Merge\"");
                    streamSecondPassMJBProcess.WriteLine("FILENAME=\"" + strMultitracOutput + "\"");
                    streamSecondPassMJBProcess.WriteLine(" ");
                }

                streamSecondPassMJBProcess.WriteLine("[PRESORT-2]");
                streamSecondPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamSecondPassMJBProcess.WriteLine("ADDRESSGROUP=\"MAIN\"");
                streamSecondPassMJBProcess.WriteLine("SELECTIVITY=NONE");
                streamSecondPassMJBProcess.WriteLine("SETTINGS=\"" + strPresortSettings + "\"");
                streamSecondPassMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                streamSecondPassMJBProcess.WriteLine("FULLSERVICEPARTICIPATION=\"FULLSERVICE\"");
                streamSecondPassMJBProcess.WriteLine(" ");

                streamSecondPassMJBProcess.WriteLine("[PRESORTEDLABELS-3]");
                streamSecondPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamSecondPassMJBProcess.WriteLine("SETTINGS=\"" + strOutputLabel + "\"");
                streamSecondPassMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                streamSecondPassMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                streamSecondPassMJBProcess.WriteLine("ABSOLUTECONTAINERNUMBERS=Y");
                streamSecondPassMJBProcess.WriteLine("FILENAME=\"" + strSortedData + "\"");
                streamSecondPassMJBProcess.WriteLine("OVERWRITE=Y");
                streamSecondPassMJBProcess.WriteLine(" ");

                streamSecondPassMJBProcess.WriteLine("[MAILDAT-4]");
                streamSecondPassMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamSecondPassMJBProcess.WriteLine("SETTINGS=\"" + strProjectName + "\"");
                streamSecondPassMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                streamSecondPassMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                streamSecondPassMJBProcess.WriteLine(" ");

                //streamSecondPassMJBProcess.WriteLine("[TERMINATE-5]");

                streamSecondPassMJBProcess.Close();
                streamSecondPassMJBProcess.Dispose();

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////////// RUNNING SECOND PASS MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdSecondPassStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cmdSecondPassStartInfo.RedirectStandardError = true;
                cmdSecondPassStartInfo.RedirectStandardOutput = true;
                cmdSecondPassStartInfo.UseShellExecute = false;
                cmdSecondPassStartInfo.CreateNoWindow = true;
                cmdSecondPassStartInfo.FileName = strBCCMailManEXE;
                cmdSecondPassStartInfo.Arguments = strCommandLine;

                cmdSecondPassProcess = Process.Start(cmdSecondPassStartInfo);
                cmdSecondPassProcess.WaitForExit();
                cmdSecondPassProcess.Dispose();

                File.Delete(strMJBProcess);
                File.Delete(strMailDatSettings);

                /*
                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////// CREATING POSTAGE STATEMENT PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                System.Environment.SetEnvironmentVariable(strEnvironmentVariable, strPostageStatementPDF, EnvironmentVariableTarget.User);

                streamPaperworkMJBProcess = new StreamWriter(strMJBProcess, false);

                streamPaperworkMJBProcess.WriteLine("[POSTAGESTATEMENT-1]");
                streamPaperworkMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamPaperworkMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                streamPaperworkMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                streamPaperworkMJBProcess.WriteLine("COMMENTS=\"" + strJobNumber + " - " + strProjectName + "\"");
                streamPaperworkMJBProcess.WriteLine("PRINTER=\"" + strPDFPrinter + "\"");
                streamPaperworkMJBProcess.WriteLine(" ");

                streamPaperworkMJBProcess.Close();
                streamPaperworkMJBProcess.Dispose();

                //////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////// RUNNING POSTAGE STATEMENT PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdPaperworkStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cmdPaperworkStartInfo.RedirectStandardError = true;
                cmdPaperworkStartInfo.RedirectStandardOutput = true;
                cmdPaperworkStartInfo.UseShellExecute = false;
                cmdPaperworkStartInfo.CreateNoWindow = true;
                cmdPaperworkStartInfo.FileName = strBCCMailManEXE;
                cmdPaperworkStartInfo.Arguments = strCommandLine;

                cmdPaperworkProcess = Process.Start(cmdPaperworkStartInfo);
                cmdPaperworkProcess.WaitForExit();

                File.Delete(strMJBProcess);

                //////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////// CREATING QUALIFICATION REPORT PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                System.Environment.SetEnvironmentVariable(strEnvironmentVariable, strQualReportPDF, EnvironmentVariableTarget.User);

                streamPaperworkMJBProcess = new StreamWriter(strMJBProcess, false);

                streamPaperworkMJBProcess.WriteLine("[QUALIFICATIONREPORT-1]");
                streamPaperworkMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamPaperworkMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                streamPaperworkMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                streamPaperworkMJBProcess.WriteLine("ABSOLUTECONTAINERNUMBERS=Y");
                streamPaperworkMJBProcess.WriteLine("PRINTER=\"" + strPDFPrinter + "\"");
                streamPaperworkMJBProcess.WriteLine(" ");

                streamPaperworkMJBProcess.Close();
                streamPaperworkMJBProcess.Dispose();

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////// RUNNING QUALIFICATION REPORT PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdPaperworkProcess = Process.Start(cmdPaperworkStartInfo);
                cmdPaperworkProcess.WaitForExit();

                File.Delete(strMJBProcess);

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////// CREATING PALLET PLACARDS PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                System.Environment.SetEnvironmentVariable(strEnvironmentVariable, strPalletPlacardsPDF, EnvironmentVariableTarget.User);

                streamPaperworkMJBProcess = new StreamWriter(strMJBProcess, false);

                streamPaperworkMJBProcess.WriteLine("[CONTAINERTAGS-1]");
                streamPaperworkMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamPaperworkMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                streamPaperworkMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                streamPaperworkMJBProcess.WriteLine("ABSOLUTECONTAINERNUMBERS=Y");
                streamPaperworkMJBProcess.WriteLine("COPIES=2");
                streamPaperworkMJBProcess.WriteLine("TAGTYPE=\"PALLET\"");
                streamPaperworkMJBProcess.WriteLine("OMITBARCODES=N");
                streamPaperworkMJBProcess.WriteLine("INFOLINE=Y");
                streamPaperworkMJBProcess.WriteLine("INFOLINEFORMAT=\"HEETER DIRECT\"");
                streamPaperworkMJBProcess.WriteLine("INFOLINEFORMAT2=\"%T    %X    %P    %N\"");
                streamPaperworkMJBProcess.WriteLine("MAILERID=\"" + strHeeterMailerID + "\"");
                streamPaperworkMJBProcess.WriteLine("NEWPAGESTREAM=Y");
                streamPaperworkMJBProcess.WriteLine("ORIENTATION=\"LANDSCAPE\"");
                streamPaperworkMJBProcess.WriteLine("MAILERNAME=\"HEETER DIRECT\"");
                streamPaperworkMJBProcess.WriteLine("MARGINLEFT=0.500");
                streamPaperworkMJBProcess.WriteLine("MARGINTOP=0.500");
                streamPaperworkMJBProcess.WriteLine("PAPER=\"Letter\"");
                streamPaperworkMJBProcess.WriteLine("PRINTER=\"" + strPDFPrinter + "\"");
                streamPaperworkMJBProcess.WriteLine(" ");

                streamPaperworkMJBProcess.Close();
                streamPaperworkMJBProcess.Dispose();

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////// RUNNING PALLET PLACARDS PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdPaperworkProcess = Process.Start(cmdPaperworkStartInfo);
                cmdPaperworkProcess.WaitForExit();

                File.Delete(strMJBProcess);

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////// CREATING TRAY TAGS PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                System.Environment.SetEnvironmentVariable(strEnvironmentVariable, strTrayTagsPDF, EnvironmentVariableTarget.User);

                streamPaperworkMJBProcess = new StreamWriter(strMJBProcess, false);

                streamPaperworkMJBProcess.WriteLine("[CONTAINERTAGS-1]");
                streamPaperworkMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                streamPaperworkMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                streamPaperworkMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                streamPaperworkMJBProcess.WriteLine("ABSOLUTECONTAINERNUMBERS=Y");
                streamPaperworkMJBProcess.WriteLine("TAGTYPE=\"IMTRAYLABEL\"");
                streamPaperworkMJBProcess.WriteLine("OMITBARCODES=N");
                streamPaperworkMJBProcess.WriteLine("ORIENTATION=\"PORTRAIT\"");
                streamPaperworkMJBProcess.WriteLine("MAILERNAME=\"HEETER DIRECT\"");
                streamPaperworkMJBProcess.WriteLine("MAILERID=\"" + strHeeterMailerID + "\"");
                streamPaperworkMJBProcess.WriteLine("NEWPAGESTREAM=Y");
                streamPaperworkMJBProcess.WriteLine("INFOLINE=Y");
                streamPaperworkMJBProcess.WriteLine("LABELWIDTH=3.250");
                streamPaperworkMJBProcess.WriteLine("LABELHEIGHT=2.000");
                streamPaperworkMJBProcess.WriteLine("ACROSS=2");
                streamPaperworkMJBProcess.WriteLine("DOWN=5");
                streamPaperworkMJBProcess.WriteLine("MARGINLEFT=1.000");
                streamPaperworkMJBProcess.WriteLine("MARGINTOP=0.625");
                streamPaperworkMJBProcess.WriteLine("PAPER=\"Letter\"");
                streamPaperworkMJBProcess.WriteLine("PRINTER=\"" + strPDFPrinter + "\"");
                streamPaperworkMJBProcess.WriteLine(" ");

                streamPaperworkMJBProcess.WriteLine("[TERMINATE-2]");

                streamPaperworkMJBProcess.Close();
                streamPaperworkMJBProcess.Dispose();

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////// RUNNING PALLET PLACARDS PDF MJB FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdPaperworkProcess = Process.Start(cmdPaperworkStartInfo);
                cmdPaperworkProcess.WaitForExit();
                cmdPaperworkProcess.Dispose();

                File.Delete(strMJBProcess);
                */

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////// UPLOADING MAIL.DAT TO MULTITRAC, IF NEEDED FOR CLIENT
                //////////////////////////////////////////////////////////////////////////////////////////
                if (bolMultitrac)
                {
                    bolMultitracSuccess = MultitracProcess("UPLOAD_MAILDAT");
                    if (!bolMultitracSuccess) return false;
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////// COPYING MAIL.DAT FILE TO MAILDOCS SHARE
                //////////////////////////////////////////////////////////////////////////////////////////
                CopyFileUsingCredentials(strDocsJobFolder + strJobNumber + "\\" + strJobNumber + ".*", strMailDatRepository,"maildoc","Heeter2013");

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////// CALCULATING TOTAL NUMBER OF MAILING RECORDS
                //////////////////////////////////////////////////////////////////////////////////////////
                aryMailingRecords = File.ReadAllLines(strSortedData);
                iMailingRecords = aryMailingRecords.Length - 1;
                strMailingRecords = iMailingRecords.ToString();
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    MultiTrac Process...    "

        static bool MultitracProcess(string strMultitracStep)
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strMultitracJobCreateXMLFile = strWorkingJobFolder + strJobNumber + " - Multitrac Job Create.xml";
            string strMultitracJobProcessXMLFile = strWorkingJobFolder + strJobNumber + " - Multitrac Job Process.xml";
            string strMultitracJobProcessBATFile = strWorkingJobFolder + strJobNumber + " - Multitrac Job Process.bat";
            string strMultitracMailDatXMLFile = strWorkingJobFolder + strJobNumber + " - Multitrac MailDat Upload.xml";
            string strMailDatFile = strDocsJobFolder + strJobNumber + Path.DirectorySeparatorChar + strJobNumber + ".zip";
            string strMultitracURL = "https://www.multitrac.com/pcws/pcode.asmx";
            double dblDaysToAddForInHome = double.Parse(strDaysToAddForInHome);
            string strInHomeDate = DateTime.Parse(strMailDate).AddDays(dblDaysToAddForInHome).ToString("yyyy-MM-dd");
            string strXMLMailDate = DateTime.Parse(strMailDate).ToString("yyyy-MM-dd");
            string strMultitracRecords;
            byte[] aryMultitracRequest;
            string strMultitracCreateJobXMLRead;
            string strMultitracUploadMailDatXMLRead;
            byte[] aryMultitracUploadMailDat;
            string strMultitracUploadMailDat;
            StreamWriter streamMultitracJobCreateXML;
            StreamWriter streamMultitracJobProcessXML;
            StreamWriter streamMultitracJobProcessBAT;
            StreamWriter streamMultitracMailDatXML;
            Stream streamMultitracJobRequest;
            Stream streamMultitracJobResponse;
            Stream streamMultitracMailDatRequest;
            Stream streamMultitracMailDatResponse;
            WebRequest webMultitracJobRequest = WebRequest.Create(strMultitracURL);
            WebResponse webMultitracJobResponse;
            WebRequest webMultitracMailDatRequest = WebRequest.Create(strMultitracURL);
            WebResponse webMultitracMailDatResponse;
            XmlTextReader xmlMultitracJobResponse;
            XmlTextReader xmlMultitracMailDatResponse;
            Process cmdMultitracJobProcess = new Process();
            ProcessStartInfo cmdMultitracJobProcessStartInfo = new ProcessStartInfo();
            
            try
            {
                switch (strMultitracStep)
                {
                    //////////////////////////////////////////////////////////////////////////////////////////
                    //////////////////////////////////////////////// CREATING AND PROCESSING THE MULTITRAC JOB
                    //////////////////////////////////////////////////////////////////////////////////////////
                    case "PROCESS_JOB":
                        // Calculating total number of mailing records.
                        strMultitracRecords = (File.ReadAllLines(strMultitracInput).Length - 1).ToString().Trim();

                        // Creating Multitrac job creation XML file.
                        streamMultitracJobCreateXML = new StreamWriter(strMultitracJobCreateXMLFile);

                        streamMultitracJobCreateXML.WriteLine(strXMLHeader);
                        streamMultitracJobCreateXML.WriteLine(strSOAPHeader);
                        streamMultitracJobCreateXML.WriteLine("<soap:Body>");
                        streamMultitracJobCreateXML.WriteLine("   <CreatePCodeFS xmlns=\"http://www.multitrac.com/\">");
                        streamMultitracJobCreateXML.WriteLine("      <userId>" + strMultitracUserName + "</userId>");
                        streamMultitracJobCreateXML.WriteLine("      <password>" + strMultitracPassword + "</password>");
                        streamMultitracJobCreateXML.WriteLine("      <jobSpecs>");
                        streamMultitracJobCreateXML.WriteLine("         <JobDetails>");
                        streamMultitracJobCreateXML.WriteLine("            <JobDet_FS>");
                        streamMultitracJobCreateXML.WriteLine("               <StartDate>" + strXMLMailDate + "</StartDate>");
                        streamMultitracJobCreateXML.WriteLine("               <InhStartDate>" + strInHomeDate + "</InhStartDate>");
                        streamMultitracJobCreateXML.WriteLine("               <DropQty>" + strMultitracRecords + "</DropQty>");
                        streamMultitracJobCreateXML.WriteLine("               <DropName>" + strProjectName + "</DropName>");
                        streamMultitracJobCreateXML.WriteLine("               <MailClass>" + ((strClassOfMail.Equals("FIRST")) ? "310" : "311") + "</MailClass>");
                        streamMultitracJobCreateXML.WriteLine("               <Version>1</Version>");
                        streamMultitracJobCreateXML.WriteLine("            </JobDet_FS>");
                        streamMultitracJobCreateXML.WriteLine("         </JobDetails>");
                        streamMultitracJobCreateXML.WriteLine("         <JobName>" + strProjectName + "</JobName>");
                        streamMultitracJobCreateXML.WriteLine("         <ExternalJobNo>" + strJobNumber + "</ExternalJobNo>");

                        if (bolRunInTestMode)
                        {
                            streamMultitracJobCreateXML.WriteLine("         <DivId>" + strMultitracTestClientID + "</DivId>");
                        }
                        else
                        {
                            streamMultitracJobCreateXML.WriteLine("         <DivId>" + strMultitracClientID + "</DivId>");
                        }

                        streamMultitracJobCreateXML.WriteLine("         <ClientId>669</ClientId>");
                        streamMultitracJobCreateXML.WriteLine("         <IMB>true</IMB>");
                        streamMultitracJobCreateXML.WriteLine("         <ByStore>false</ByStore>");
                        streamMultitracJobCreateXML.WriteLine("         <EntryPoint>15290</EntryPoint>");
                        streamMultitracJobCreateXML.WriteLine("         <ACS>N</ACS>");
                        streamMultitracJobCreateXML.WriteLine("         <ServiceOption>F</ServiceOption>");
                        streamMultitracJobCreateXML.WriteLine("         <OEL>true</OEL>");
                        streamMultitracJobCreateXML.WriteLine("      </jobSpecs>");
                        streamMultitracJobCreateXML.WriteLine("   </CreatePCodeFS>");
                        streamMultitracJobCreateXML.WriteLine("</soap:Body>");
                        streamMultitracJobCreateXML.WriteLine("</soap:Envelope>");

                        streamMultitracJobCreateXML.Close();
                        streamMultitracJobCreateXML.Dispose();

                        // Posting create job XML file to Multitrac and extracting job number from response.
                        strMultitracCreateJobXMLRead = new StreamReader(strMultitracJobCreateXMLFile).ReadToEnd();

                        webMultitracJobRequest.ContentType = "text/xml;charset=utf-8";
                        webMultitracJobRequest.Method = "POST";

                        aryMultitracRequest = Encoding.ASCII.GetBytes(strMultitracCreateJobXMLRead);
                        webMultitracJobRequest.ContentLength = aryMultitracRequest.Length;

                        streamMultitracJobRequest = webMultitracJobRequest.GetRequestStream();
                        streamMultitracJobRequest.Write(aryMultitracRequest, 0, aryMultitracRequest.Length);

                        webMultitracJobResponse = webMultitracJobRequest.GetResponse();
                        streamMultitracJobResponse = webMultitracJobResponse.GetResponseStream();

                        xmlMultitracJobResponse = new XmlTextReader(streamMultitracJobResponse);

                        while (xmlMultitracJobResponse.Read())
                        {
                            if (xmlMultitracJobResponse.Name.Equals("JobNo"))
                            {
                                strMultitracJobNumber = xmlMultitracJobResponse.ReadString();
                            }
                        }

                        streamMultitracJobRequest.Close();
                        streamMultitracJobRequest.Dispose();
                        streamMultitracJobResponse.Close();
                        streamMultitracJobResponse.Dispose();

                        // Creating Multitrac job process XML file.
                        streamMultitracJobProcessXML = new StreamWriter(strMultitracJobProcessXMLFile);

                        streamMultitracJobProcessXML.WriteLine(strXMLHeader);
                        streamMultitracJobProcessXML.WriteLine("<parameters>");

                        // Processing information.
                        streamMultitracJobProcessXML.WriteLine("   <LoginID>" + strMultitracUserName + "</LoginID>");
                        streamMultitracJobProcessXML.WriteLine("   <LoginPSW>" + strMultitracPassword + "</LoginPSW>");
                        streamMultitracJobProcessXML.WriteLine("   <JobNo>" + strMultitracJobNumber + "</JobNo>");
                        streamMultitracJobProcessXML.WriteLine("   <InputFile>" + strMultitracInput + "</InputFile>");
                        streamMultitracJobProcessXML.WriteLine("   <OutputFile>" + strMultitracOutput + "</OutputFile>");

                        // File information.
                        streamMultitracJobProcessXML.WriteLine("   <FileType>CSV</FileType>");
                        streamMultitracJobProcessXML.WriteLine("   <ClientIDFieldName>CLIENT ID</ClientIDFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <NameFieldName>FULLNAME</NameFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <Address1FieldName>ADDRESS 1</Address1FieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <Address2FieldName>ADDRESS 2</Address2FieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <CityFieldName>CITY</CityFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <StateFieldName>STATE</StateFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <ZipFieldName>ZIPCODE</ZipFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <BarcodeFieldName>ZIP+4-DP COMBO</BarcodeFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <VersionFieldName>VERSION NUMBER</VersionFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <PhoneFieldName>PHONE NUMBER</PhoneFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <EmailFieldName>EMAIL</EmailFieldName>");
                        streamMultitracJobProcessXML.WriteLine("   <SeedFlagFieldName>SEED</SeedFlagFieldName>");

                        // Unused elements.
                        streamMultitracJobProcessXML.WriteLine("   <ClientIDStart />");
                        streamMultitracJobProcessXML.WriteLine("   <ClientIDLength />");
                        streamMultitracJobProcessXML.WriteLine("   <NameStart />");
                        streamMultitracJobProcessXML.WriteLine("   <NameLength />");
                        streamMultitracJobProcessXML.WriteLine("   <Address1Start />");
                        streamMultitracJobProcessXML.WriteLine("   <Address1Length />");
                        streamMultitracJobProcessXML.WriteLine("   <Address2Start />");
                        streamMultitracJobProcessXML.WriteLine("   <Address2Length />");
                        streamMultitracJobProcessXML.WriteLine("   <CityStart />");
                        streamMultitracJobProcessXML.WriteLine("   <CityLength />");
                        streamMultitracJobProcessXML.WriteLine("   <StateStart />");
                        streamMultitracJobProcessXML.WriteLine("   <StateLength />");
                        streamMultitracJobProcessXML.WriteLine("   <ZipStart />");
                        streamMultitracJobProcessXML.WriteLine("   <ZipLength />");
                        streamMultitracJobProcessXML.WriteLine("   <Zip4FieldName />");
                        streamMultitracJobProcessXML.WriteLine("   <Zip4Start />");
                        streamMultitracJobProcessXML.WriteLine("   <Zip4Length />");
                        streamMultitracJobProcessXML.WriteLine("   <DPBCFieldName />");
                        streamMultitracJobProcessXML.WriteLine("   <DPBCStart />");
                        streamMultitracJobProcessXML.WriteLine("   <DPBCLength />");
                        streamMultitracJobProcessXML.WriteLine("   <BarcodeStart />");
                        streamMultitracJobProcessXML.WriteLine("   <BarcodeLength />");
                        streamMultitracJobProcessXML.WriteLine("   <VersionStart />");
                        streamMultitracJobProcessXML.WriteLine("   <VersionLength />");
                        streamMultitracJobProcessXML.WriteLine("   <PhoneStart />");
                        streamMultitracJobProcessXML.WriteLine("   <PhoneLength />");
                        streamMultitracJobProcessXML.WriteLine("   <EmailStart />");
                        streamMultitracJobProcessXML.WriteLine("   <EmailLength />");
                        streamMultitracJobProcessXML.WriteLine("   <StoreFieldName />");
                        streamMultitracJobProcessXML.WriteLine("   <StoreStart />");
                        streamMultitracJobProcessXML.WriteLine("   <StoreLength />");
                        streamMultitracJobProcessXML.WriteLine("   <OptEndrsFieldName />");
                        streamMultitracJobProcessXML.WriteLine("   <OptEndrsStart />");
                        streamMultitracJobProcessXML.WriteLine("   <OptEndrsLength />");
                        streamMultitracJobProcessXML.WriteLine("   <SeedFlagStart />");
                        streamMultitracJobProcessXML.WriteLine("   <SeedFlagLength />");
                        streamMultitracJobProcessXML.WriteLine("   <EntryFieldName />");
                        streamMultitracJobProcessXML.WriteLine("   <EntryStart />");
                        streamMultitracJobProcessXML.WriteLine("   <EntryLength />");
                        streamMultitracJobProcessXML.WriteLine("   <PRateFieldName />");
                        streamMultitracJobProcessXML.WriteLine("   <PRateStart />");
                        streamMultitracJobProcessXML.WriteLine("   <PRateLength />");
                        streamMultitracJobProcessXML.WriteLine("   <PCStartChar />");
                        streamMultitracJobProcessXML.WriteLine("   <PCEndChar />");
                        streamMultitracJobProcessXML.WriteLine("</parameters>");

                        streamMultitracJobProcessXML.Close();
                        streamMultitracJobProcessXML.Dispose();

                        // Processing the Multitrac job from the command line.
                        streamMultitracJobProcessBAT = new StreamWriter(strMultitracJobProcessBATFile);
                        streamMultitracJobProcessBAT.WriteLine("\"" + strMultitracEXE + "\" \"" + strMultitracJobProcessXMLFile + "\"");
                        
                        streamMultitracJobProcessBAT.Close();
                        streamMultitracJobProcessBAT.Dispose();

                        cmdMultitracJobProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        cmdMultitracJobProcessStartInfo.RedirectStandardError = false;
                        cmdMultitracJobProcessStartInfo.RedirectStandardOutput = false;
                        cmdMultitracJobProcessStartInfo.UseShellExecute = true;
                        cmdMultitracJobProcessStartInfo.CreateNoWindow = true;
                        cmdMultitracJobProcessStartInfo.FileName = "CMD.exe";
                        cmdMultitracJobProcessStartInfo.Arguments = "/C \"" + strMultitracJobProcessBATFile + "\"";

                        cmdMultitracJobProcess = Process.Start(cmdMultitracJobProcessStartInfo);
                        cmdMultitracJobProcess.WaitForExit();

                        File.Delete(strMultitracJobProcessBATFile);

                        break;

                    //////////////////////////////////////////////////////////////////////////////////////////
                    ////////////////////////////////////////////////////// UPLOADING THE MAIL.DAT TO MULTITRAC
                    //////////////////////////////////////////////////////////////////////////////////////////
                    case "UPLOAD_MAILDAT":
                        // Encoding the Mail.dat with Base 64 Binary.
                        aryMultitracUploadMailDat = File.ReadAllBytes(strMailDatFile);
                        strMultitracUploadMailDat = Convert.ToBase64String(aryMultitracUploadMailDat);

                        // Creating the Mail.dat upload XML file.
                        streamMultitracMailDatXML = new StreamWriter(strMultitracMailDatXMLFile);

                        streamMultitracMailDatXML.WriteLine(strXMLHeader);
                        streamMultitracMailDatXML.WriteLine(strSOAPHeader);
                        streamMultitracMailDatXML.WriteLine("<soap:Body>");
                        streamMultitracMailDatXML.WriteLine("   <LoadJobMailDatFile xmlns=\"http://www.multitrac.com/\">");
                        streamMultitracMailDatXML.WriteLine("      <userId>" + strMultitracUserName + "</userId>");
                        streamMultitracMailDatXML.WriteLine("      <password>" + strMultitracPassword + "</password>");
                        streamMultitracMailDatXML.WriteLine("      <jobid>" + strMultitracJobNumber + "</jobid>");
                        streamMultitracMailDatXML.WriteLine("      <filename>" + strMailDatFile + "</filename>");
                        streamMultitracMailDatXML.WriteLine("      <file>" + strMultitracUploadMailDat + "</file>");
                        streamMultitracMailDatXML.WriteLine("   </LoadJobMailDatFile>");
                        streamMultitracMailDatXML.WriteLine("</soap:Body>");
                        streamMultitracMailDatXML.WriteLine("</soap:Envelope>");

                        streamMultitracMailDatXML.Close();
                        streamMultitracMailDatXML.Dispose();

                        // Posting Mail.dat upload XML file to Multitrac.
                        strMultitracUploadMailDatXMLRead = new StreamReader(strMultitracMailDatXMLFile).ReadToEnd();

                        webMultitracMailDatRequest.ContentType = "text/xml;charset=utf-8";
                        webMultitracMailDatRequest.Method = "POST";

                        aryMultitracRequest = Encoding.ASCII.GetBytes(strMultitracUploadMailDatXMLRead);
                        webMultitracMailDatRequest.ContentLength = aryMultitracRequest.Length;

                        streamMultitracMailDatRequest = webMultitracMailDatRequest.GetRequestStream();
                        streamMultitracMailDatRequest.Write(aryMultitracRequest, 0, aryMultitracRequest.Length);

                        webMultitracMailDatResponse = webMultitracMailDatRequest.GetResponse();
                        streamMultitracMailDatResponse = webMultitracMailDatResponse.GetResponseStream();

                        xmlMultitracMailDatResponse = new XmlTextReader(streamMultitracMailDatResponse);

                        streamMultitracMailDatRequest.Close();
                        streamMultitracMailDatRequest.Dispose();
                        streamMultitracMailDatResponse.Close();
                        streamMultitracMailDatResponse.Dispose();

                        break;

                    default:
                        break;
                }
            }
            //WebException - SystemException - SecurityException - IOException
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }
            finally
            {
                cmdMultitracJobProcess.Dispose();
            }

            return true;
        }

        #endregion


        #region "    Payment Information...    "

        static bool PaymentInformation()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strPostageStatementPRN = strWorkingJobFolder + strJobNumber + " - Postage Statement.prn";
            string strMJBProcess = strTaskmasterJobsFolder + strProjectName + ".mjb";
            string strCommandLine = "-j \"" + Path.GetFileName(strMJBProcess) + "\" -u \"AUTO\" -w \"AUTO\" -r";
            string strPostagePRNLine;
            int iPostageStartIndex = 0;
            int iPostageSubstringLength = 0;
            string[] aryUserInfoTableLine;
            string strUserInfoTableLine;
            bool bolFoundUserID = false;
            string strUserInfoID;
            Process cmdPostageProcess = new Process();
            ProcessStartInfo cmdPostageStartInfo = new ProcessStartInfo();
            StreamWriter streamPostageMJBProcess;
            StreamReader streamUserInfoTableFile = new StreamReader(strUserInfoTable);
            
            try
            {   
                try
                {
                    //////////////////////////////////////////////////////////////////////////////////////////
                    ////////////////////////////////////////// COLLECTING THE POSTAGE AMOUNT, IF PRINTMAIL JOB
                    //////////////////////////////////////////////////////////////////////////////////////////
                    if (strOutputType.Equals("PRINTMAIL"))
                    {
                        streamPostageMJBProcess = new StreamWriter(strMJBProcess, false);

                        streamPostageMJBProcess.WriteLine("[POSTAGESTATEMENT-1]");
                        streamPostageMJBProcess.WriteLine("LIST=\"" + strBCCDatabase + "\"");
                        streamPostageMJBProcess.WriteLine("PRESORTNAME=\"" + strProjectName + "\"");
                        streamPostageMJBProcess.WriteLine("STREAMLIST=\"MERGED;AUTO/NONAUTO;AUTO;MACH;SINGLE PC\"");
                        streamPostageMJBProcess.WriteLine("FILENAME=\"" + strPostageStatementPRN + "\"");
                        streamPostageMJBProcess.WriteLine("PRINTER=\"" + strTextPrinter + "\"");
                        streamPostageMJBProcess.WriteLine("PRINTTOFILE=Y");
                        streamPostageMJBProcess.WriteLine("PRINTMODE=\"CONSOLIDATED\"");
                        streamPostageMJBProcess.WriteLine(" ");

                        streamPostageMJBProcess.WriteLine("[TERMINATE-2]");

                        streamPostageMJBProcess.Close();
                        streamPostageMJBProcess.Dispose();

                        //////////////////////////////////////////////////////////////////////////////////////////
                        /////////////////////////////////////////////////// RUNNING POSTAGE STATEMENT PRN MJB FILE
                        //////////////////////////////////////////////////////////////////////////////////////////
                        cmdPostageStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        cmdPostageStartInfo.RedirectStandardError = true;
                        cmdPostageStartInfo.RedirectStandardOutput = true;
                        cmdPostageStartInfo.UseShellExecute = false;
                        cmdPostageStartInfo.CreateNoWindow = true;
                        cmdPostageStartInfo.FileName = strBCCMailManEXE;
                        cmdPostageStartInfo.Arguments = strCommandLine;

                        cmdPostageProcess = Process.Start(cmdPostageStartInfo);
                        cmdPostageProcess.WaitForExit();

                        File.Delete(strMJBProcess);

                        //////////////////////////////////////////////////////////////////////////////////////////
                        /////////////////////////// READING THE POSTAGE AMOUNT FROM THE POSTAGE STATEMENT PRN FILE
                        //////////////////////////////////////////////////////////////////////////////////////////
                        StreamReader streamPostagePRNFile = new StreamReader(strPostageStatementPRN);

                        while (!streamPostagePRNFile.EndOfStream)
                        {
                            strPostagePRNLine = streamPostagePRNFile.ReadLine();

                            if (strPostagePRNLine.ToLower().Contains("(add parts totals)"))
                            {
                                iPostageStartIndex = strPostagePRNLine.ToLower().LastIndexOf("(add parts totals)") + 18;
                                iPostageSubstringLength = strPostagePRNLine.Length - iPostageStartIndex;

                                dblPostageAmount = double.Parse(strPostagePRNLine.Substring(iPostageStartIndex, iPostageSubstringLength));

                                break;
                            }
                        }

                        streamPostagePRNFile.Close();
                        streamPostagePRNFile.Dispose();
                    }
                    else
                    {
                        dblPostageAmount = 0.0;
                    }
                }
                catch (Exception exception)
                {
                    LogFile(exception.ToString(), true);
                    return false;
                }
                finally
                {
                    cmdPostageProcess.Close();
                    cmdPostageProcess.Dispose();
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////// DETERMINING IF USER IS REQUIRED TO PAY BY CREDIT CARD
                //////////////////////////////////////////////////////////////////////////////////////////
                try
                {
                    while (!streamUserInfoTableFile.EndOfStream)
                    {
                        strUserInfoTableLine = streamUserInfoTableFile.ReadLine();
                        aryUserInfoTableLine = strUserInfoTableLine.Split(',');

                        strUserInfoID = aryUserInfoTableLine[0].ToString().ToUpper().Trim();

                        if (strLoginID == strUserInfoID)
                        {
                            bolFoundUserID = true;
                            strCreditCardRequired = aryUserInfoTableLine[7].ToString().ToUpper().Trim();

                            break;
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
                    streamUserInfoTableFile.Close();
                    streamUserInfoTableFile.Dispose();
                }

                if (!bolFoundUserID)
                {
                    LogFile("A login ID of '" + strLoginID + "' could not be found in the user matrix :  " + strUserInfoTable, true);
                    return false;
                }
                else
                {
                    if (!(strCreditCardRequired.Equals("NONE")) && !(strCreditCardRequired.Equals("ALL")) && !(strCreditCardRequired.Equals("POSTAGE")))
                    {
                        LogFile("An incorrect value of '" + strCreditCardRequired + "' was used in the user matrix :  " + strUserInfoTable, true);
                        return false;
                    }
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////// DETERMINING JOB STATUS BASED ON OUTPUT TYPE AND USER TYPE
                //////////////////////////////////////////////////////////////////////////////////////////
                try
                {
                    if (!bolPaymentExpected)
                    {
                        strJobStatus = "READY TO PROCESS";
                    }
                    else
                    {
                        switch (strOutputType)
                        {
                            case "EMAILPDF":
                                strJobStatus = "READY TO PROCESS";
                                break;

                            case "PICKANDPACK":
                                strJobStatus = "PROCESSED";
                                break;

                            default:
                                if (strCreditCardRequired.Equals("NONE"))
                                {
                                    strJobStatus = "READY TO PROCESS";
                                }
                                else
                                {
                                    strJobStatus = "PAYMENT NEEDED";
                                }

                                break;
                        }
                    }
                }
                catch (Exception exception)
                {
                    LogFile(exception.ToString(), true);
                    return false;
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    Direct Response Process...    "

        static bool DirectResponseProcess()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////// OBTAINING KIT KEY VALUE FROM XML FILE
            //////////////////////////////////////////////////////////////////////////////////////////
            string strKitKeyValue = "";

            try
            {
                if (!string.IsNullOrEmpty(strAdditionalKitKeyField))
                {
                    using (XmlTextReader xmlItemDataXMLStream = new XmlTextReader(strItemDataXMLFullName))
                    {
                        while (xmlItemDataXMLStream.Read())
                        {
                            if (xmlItemDataXMLStream.Name == strAdditionalKitKeyField)
                            {
                                strKitKeyValue = xmlItemDataXMLStream.ReadString();
                            }
                        }
                    }
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                ///////////////////////////////// MAKING SURE KIT KEY VALUE ISN'T LONGER THAN 9 CHARACTERS
                //////////////////////////////////////////////////////////////////////////////////////////
                if (strKitKeyValue.Length > 9)
                {
                    LogFile("The contents of the specified kit key field - " + strAdditionalKitKeyField + " - must be no more than 9 characters.", true);
                    return false;
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
                //////////////////////////////////////////////////////////////////////////////////////////
                string strDateTimeStamp = DateTime.Now.Date.ToString("MMddyyyy") + DateTime.Now.Hour.ToString().PadLeft(2,'0') + DateTime.Now.Minute.ToString().PadLeft(2,'0') + DateTime.Now.Second.ToString().PadLeft(2,'0');
                strDirectResponseCSV = strWorkingJobFolder + strJobNumber + " - DR Upload_" + strDateTimeStamp + ".csv";
                strDirectResponseItemCode = strPrimaryKitKey + (string.IsNullOrEmpty(strKitKeyValue) ? "" : "_" + strKitKeyValue);
                string strDirectResponseShipDate = "";

                if (string.IsNullOrEmpty(strMailDate))
                {
                    if (string.IsNullOrEmpty(strDeliveryDate))
                    {
                        if (strOutputType.Equals("EMAILPDF"))
                        {
                            strDirectResponseShipDate = DateTime.Now.Date.ToString("MM/dd/yyyy");
                        }
                        else
                        {
                            LogFile("Neither a Mail Date nor a Delivery Date were specified in the Item Data XML file.", true);
                            return false;
                        }

                    }
                    else
                    {
                        strDirectResponseShipDate = DateTime.Parse(strDeliveryDate).ToString("MM/dd/yyyy");
                    }
                }
                else
                {
                    strDirectResponseShipDate = strMailDate;
                }

                using (StreamWriter streamDirectResponseCSVWrite = new StreamWriter(strDirectResponseCSV))
                {
                    //////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////// CREATING THE DIRECT RESPONSE UPLOAD CSV FILE
                    //////////////////////////////////////////////////////////////////////////////////////////
                    streamDirectResponseCSVWrite.WriteLine("\"INVENTORY #\"," +
                                                           "\"BRANCH #\"," +
                                                           "\"DIVISION\"," +
                                                           "\"COSTCENTER\"," +
                                                           "\"SERVICE TYPE\"," +
                                                           "\"CARRIER\"," +
                                                           "\"REQUESTED SHIPDATE\"," +
                                                           "\"FULL NAME\"," +
                                                           "\"COMPANY NAME\"," +
                                                           "\"ADDRESS 1\"," +
                                                           "\"ADDRESS 2\"," +
                                                           "\"CITY\"," +
                                                           "\"STATE\"," +
                                                           "\"COUNTRY\"," +
                                                           "\"ZIP CODE\"," +
                                                           "\"EMAIL\"," +
                                                           "\"PHONE NUMBER\"," +
                                                           "\"COMMENT\"," +
                                                           "\"EXTERNAL ORDER #\"," +
                                                           "\"CONTROL NUMBER\"," +
                                                           "\"SUPPRESS PRINTING?\"," +
                                                           "\"CUSTOM_FIELD PROGRAMID%VALUE%CUSTOM_FIELD PROGRAMID%VALUE|\"," +
                                                           "\"ITEM #\"," +
                                                           "\"ITEM REVISION\"," +
                                                           "\"QTY\"");

                    streamDirectResponseCSVWrite.WriteLine("\"" + strDRInventoryName + "\"," +
                                                           "\"" + strCreditCardRequired + "\"," +
                                                           "\"\"," +
                                                           "\"\"," +
                                                           "\"Standard\"," +
                                                           "\"" + (strOutputType.Equals("PRINTMAIL") ? "26" : strOutputType.Equals("EMAILPDF") ? "36" : "9") + "\"," +
                                                           "\"" + strDirectResponseShipDate + "\"," +
                                                           "\"" + (string.IsNullOrEmpty(strShipName) ? "GENERAL MAIL FACILITY" : strShipName) + "\"," +
                                                           "\"" + (strOutputType.Equals("PRINTMAIL") ? "USPS" : (string.IsNullOrEmpty(strShipCareOf) ? strCompanyName : strShipCareOf)) + "\"," +
                                                           "\"" + (string.IsNullOrEmpty(strShipAdd1) ? "1000 CALIFORNIA AVENUE" : strShipAdd1) + "\"," +
                                                           "\"" + (string.IsNullOrEmpty(strShipAdd2) ? "" : strShipAdd2) + "\"," +
                                                           "\"" + (string.IsNullOrEmpty(strShipCity) ? "PITTSBURGH" : strShipCity) + "\"," +
                                                           "\"" + (string.IsNullOrEmpty(strShipState) ? "PA" : strShipState) + "\"," +
                                                           "\"USA\"," +
                                                           "\"" + (string.IsNullOrEmpty(strShipZip) ? "15290" : strShipZip) + "\"," +
                                                           "\"" + (string.IsNullOrEmpty(strDeliveryEmail) ? (string.IsNullOrEmpty(strUserEmail) ? strInternalEmailTo : strUserEmail) : strDeliveryEmail) + "\"," +
                                                           "\"999-999-9999\"," +
                                                           "\"" + (string.IsNullOrEmpty(strProductionPDFFile) ? "" : Path.GetFileName(strProductionPDFFile)) + "\"," +
                                                           "\"" + strDirectResponseOrderNumber + "\"," +
                                                           "\"\"," +
                                                           "\"0\"," +
                                                           "\"\"," +
                                                           "\"" + (strOutputType.Equals("EMAILPDF") ? "EMAIL" : strDirectResponseItemCode) + "\"," +
                                                           "\"none\"," +
                                                           "\"" + (iMailingRecords.Equals(0) ? iInputQuantity.ToString() : iMailingRecords.ToString()) + "\"");
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////// UPLOADING THE DIRECT RESPONSE CSV TO FTP SITE
                //////////////////////////////////////////////////////////////////////////////////////////
                byte[] aryDirectResponseCSV;

                using (StreamReader streamDirectResponseCSVRead = new StreamReader(strDirectResponseCSV))
                {
                    aryDirectResponseCSV = Encoding.ASCII.GetBytes(streamDirectResponseCSVRead.ReadToEnd());
                }

                FtpWebRequest ftpCSVUploadRequest = (FtpWebRequest)FtpWebRequest.Create(strDirectResponseFTP + Path.GetFileName(strDirectResponseCSV));

                ftpCSVUploadRequest.Credentials = new NetworkCredential(strDirectResponseUserName, strDirectResponsePassword);
                ftpCSVUploadRequest.ContentLength = aryDirectResponseCSV.Length;
                ftpCSVUploadRequest.Method = WebRequestMethods.Ftp.UploadFile;

                using (Stream streamFTPRequest = ftpCSVUploadRequest.GetRequestStream())
                {
                    streamFTPRequest.Write(aryDirectResponseCSV, 0, aryDirectResponseCSV.Length);
                }
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    GMC Process...    "

        static bool GMCProcess()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            
            // PDF names
            strProductionPDFFile = strProductionJobFolder + strJobNumber + " - " + strProjectName + " %j.pdf";
            strApogeePDFFile = strApogeeClientFolder + strJobNumber + " - " + strProjectName + "_APOGEE.pdf";
            strWatermarkPDFFile = strCSRJobFolder + strJobNumber + " - " + strProjectName + "_PROOF.pdf";

            // Modifying WFD Name for parameters
            strWFDFullName = "\"" + strWFDFullName + "\" ";
            
            // Job configuration file paramaters
            string strDataInputParameter = " -dif" + strDataInputModule + " \"" + strSortedData + "\" ";
            string strXMLInputParameter = " -dif" + strXMLInputModule + " \"" + strItemDataXMLFullName + "\" ";
            string strOutputEngineParameter = " -e PDF ";
            string strProductionPDFFileParameter = " -f \"" + strProductionPDFFile + "\" ";
            string strApogeePDFFileParameter = " -f \"" + strApogeePDFFile + "\" ";
            string strWatermarkPDFFileParameter = " -f \"" + strWatermarkPDFFile + "\" ";
            string strDataGeneratorMultiParameter = " -dgfrom" + strDataGeneratorModule + " 1 -dgto" + strDataGeneratorModule + " " + iInputQuantity.ToString().Trim() + " ";
            string strDataGeneratorSingleParameter = " -dgfrom" + strDataGeneratorModule + " 1 -dgto" + strDataGeneratorModule + " 1 ";
            string strLowResConfigParameter = " -c \"" + strLowResConfigFile + "\" ";
            string strHighResConfigParameter = " -c \"" + strHighResConfigFile + "\" ";
            string strGMCErrorsLogParameter = " -la \"" + strLogFile + "\" -logtime -retadv2 ";
            string strFirstRecordParameter = " -diUseRangeForProduction" + strDataInputModule + " true -diStartRecord" + strDataInputModule + " 2 -diEndRecord" + strDataInputModule + " 2 ";
            string strProductionParameter = " -" + strProductionParameterName + strParamInputModule + " true ";
            string strStorefrontIDParameter = " -" + strStorefrontIDParameterName + strParamInputModule + " \"" + strDirectResponseOrderNumber + "\" ";
            string strImageParameter = "";
            string strSplitByGroupParameter = " -splitbygroup ";
            string strDigitalOutputModuleParameter = " -o " + strDigitalOutputModule + " ";
            string strOffsetOutputModuleParameter = " -o " + strOffsetOutputModule + " ";
            string strWatermarkOutputModuleParameter = " -o " + strWatermarkOutputModule + " ";
            string strDefaultOutputModuleParameter = " -o " + strDefaultOutputModule + " ";

            int iCount = 0;

            if (!lstImageFiles.Count.Equals(0))
            {
                foreach (string strImageName in lstImageFiles)
                {
                    iCount++;

                    strImageParameter = strImageParameter + " -" + strImageParameterRootName + iCount.ToString().Trim() + strImageParamInputModule + " \"" + strPNetImagesJobFolder + strImageName + "\" ";
                }
            }

            // GMC command line needs
            Process cmdGMCProcess = new Process();
            ProcessStartInfo cmdGMCProcessStartInfo = new ProcessStartInfo();
            string strGMCJobBatchFile = strWorkingJobFolder + strJobNumber + " - " + strProjectName + ".txt";
            StreamWriter streamGMCJobBatchFile = new StreamWriter(strGMCJobBatchFile);
            string strCommandLine = " -s " + strGMCServerIPAddress + " -p " + strGMCServerPort + " -multiple \"" + strGMCJobBatchFile + "\" -retadv -la \"" + strLogFile + "\"";
            int iGMCExitCode = 0;

            // PDF Background ON or OFF
            string strPrintGroup1TrueParameter = " -pug1 true ";
            string strPrintGroup1FalseParameter = " -pug1 false ";

            // PDF Variable ON or OFF
            string strPrintGroup2TrueParameter = " -pug2 true ";
            string strPrintGroup2FalseParameter = " -pug2 false ";

            // PDF Watermark ON or OFF
            string strPrintGroup3TrueParameter = " -pug3 true ";
            string strPrintGroup3FalseParameter = " -pug3 false ";
            
            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////// CREATING THE JOB BATCH FILE
                //////////////////////////////////////////////////////////////////////////////////////////
                switch (strOutputType)
                {
                    case "PRINTSHIP":
                        // Production PDF
                        if (iInputQuantity > iDataSplitValue)
                        {
                            // Offset PDF
                            bolPassToApogee = true;

                            streamGMCJobBatchFile.WriteLine(strWFDFullName + 
                                                            strXMLInputParameter +
                                                            strDataGeneratorMultiParameter + 
                                                            strProductionParameter + 
                                                            strStorefrontIDParameter +
                                                            strOffsetOutputModuleParameter +
                                                            strOutputEngineParameter + 
                                                            strApogeePDFFileParameter +
                                                            strPrintGroup1TrueParameter + 
                                                            strPrintGroup2TrueParameter +
                                                            strPrintGroup3FalseParameter +
                                                            strImageParameter +
                                                            strGMCErrorsLogParameter);
                        }
                        else
                        {
                            // HP PDF
                            streamGMCJobBatchFile.WriteLine(strWFDFullName +
                                                            strXMLInputParameter +
                                                            strDataGeneratorMultiParameter + 
                                                            strProductionParameter +
                                                            strStorefrontIDParameter +
                                                            strDigitalOutputModuleParameter +
                                                            strOutputEngineParameter + 
                                                            strProductionPDFFileParameter +
                                                            strSplitByGroupParameter +
                                                            strPrintGroup1TrueParameter + 
                                                            strPrintGroup2TrueParameter +
                                                            strPrintGroup3FalseParameter +
                                                            strImageParameter +
                                                            strGMCErrorsLogParameter);
                        }

                        // Proof PDF
                        streamGMCJobBatchFile.WriteLine(strWFDFullName +
                                                        strXMLInputParameter +
                                                        strDataGeneratorSingleParameter + 
                                                        strProductionParameter +
                                                        strStorefrontIDParameter +
                                                        strWatermarkOutputModuleParameter +
                                                        strOutputEngineParameter + 
                                                        strWatermarkPDFFileParameter +
                                                        strPrintGroup1TrueParameter + 
                                                        strPrintGroup2TrueParameter +
                                                        strPrintGroup3TrueParameter +
                                                        strImageParameter +
                                                        strGMCErrorsLogParameter);
                        break;

                    case "PRINTMAIL":
                        // Production PDF
                        if ((strSplitOverride.Equals("DIGI")) || ((iMailingRecords > iDataSplitValue) && (!strSplitOverride.Equals("HP"))))
                        {
                            // Offset PDF
                            bolPassToApogee = true;

                            streamGMCJobBatchFile.WriteLine(strWFDFullName +
                                                            strXMLInputParameter +
                                                            strDataInputParameter +
                                                            strFirstRecordParameter +
                                                            strProductionParameter +
                                                            strStorefrontIDParameter +
                                                            strOffsetOutputModuleParameter +
                                                            strOutputEngineParameter +
                                                            strApogeePDFFileParameter +
                                                            strPrintGroup1TrueParameter +
                                                            strPrintGroup2FalseParameter +
                                                            strPrintGroup3FalseParameter +
                                                            strImageParameter +
                                                            strGMCErrorsLogParameter);

                            // Digimaster PDF
                            streamGMCJobBatchFile.WriteLine(strWFDFullName +
                                                            strXMLInputParameter +
                                                            strDataInputParameter +
                                                            strProductionParameter +
                                                            strStorefrontIDParameter +
                                                            strDigitalOutputModuleParameter +
                                                            strOutputEngineParameter +
                                                            strProductionPDFFileParameter +
                                                            strPrintGroup1FalseParameter +
                                                            strPrintGroup2TrueParameter +
                                                            strPrintGroup3FalseParameter +
                                                            strSplitByGroupParameter +
                                                            strImageParameter +
                                                            strGMCErrorsLogParameter);
                        }
                        else
                        {
                            //  HP PDF
                            streamGMCJobBatchFile.WriteLine(strWFDFullName +
                                                            strXMLInputParameter +
                                                            strDataInputParameter +
                                                            strProductionParameter +
                                                            strStorefrontIDParameter +
                                                            strDigitalOutputModuleParameter +
                                                            strOutputEngineParameter +
                                                            strProductionPDFFileParameter +
                                                            strPrintGroup1TrueParameter +
                                                            strPrintGroup2TrueParameter +
                                                            strPrintGroup3FalseParameter +
                                                            strSplitByGroupParameter +
                                                            strImageParameter +
                                                            strGMCErrorsLogParameter);
                        }

                        // Proof PDF
                        streamGMCJobBatchFile.WriteLine(strWFDFullName +
                                                        strXMLInputParameter +
                                                        strDataInputParameter + 
                                                        strFirstRecordParameter + 
                                                        strProductionParameter +
                                                        strStorefrontIDParameter +
                                                        strWatermarkOutputModuleParameter +
                                                        strOutputEngineParameter + 
                                                        strWatermarkPDFFileParameter +
                                                        strPrintGroup1TrueParameter +
                                                        strPrintGroup2TrueParameter +
                                                        strPrintGroup3TrueParameter +
                                                        strImageParameter +
                                                        strGMCErrorsLogParameter);
                        break;

                    case "EMAILPDF":
                        // HP (i.e. full-digital) PDF
                        streamGMCJobBatchFile.WriteLine(strWFDFullName +
                                                        strXMLInputParameter +
                                                        strDataGeneratorSingleParameter + 
                                                        strProductionParameter +
                                                        strStorefrontIDParameter +
                                                        strDefaultOutputModuleParameter +
                                                        strOutputEngineParameter + 
                                                        strProductionPDFFileParameter.Replace(" %j","") +
                                                        strPrintGroup1TrueParameter +
                                                        strPrintGroup2TrueParameter +
                                                        strPrintGroup3FalseParameter +
                                                        strImageParameter +
                                                        strGMCErrorsLogParameter);
                        break;

                    default:
                        break;
                }

                streamGMCJobBatchFile.Close();

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////// PROCESSING GMC JOB BATCH FILE FROM COMMAND LINE
                //////////////////////////////////////////////////////////////////////////////////////////
                cmdGMCProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                cmdGMCProcessStartInfo.RedirectStandardError = false;
                cmdGMCProcessStartInfo.RedirectStandardOutput = false;
                cmdGMCProcessStartInfo.UseShellExecute = true;
                cmdGMCProcessStartInfo.CreateNoWindow = true;
                cmdGMCProcessStartInfo.FileName = strGMCConsoleExe;
                cmdGMCProcessStartInfo.Arguments = strCommandLine;
                cmdGMCProcessStartInfo.Verb = "runas";

                cmdGMCProcess = Process.Start(cmdGMCProcessStartInfo);
                cmdGMCProcess.WaitForExit();

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////// HANDLING GMC SUCCESS OR FAILURE
                //////////////////////////////////////////////////////////////////////////////////////////
                iGMCExitCode = cmdGMCProcess.ExitCode;
                
                if (!iGMCExitCode.Equals(0) && !iGMCExitCode.Equals(20))
                {
                    LogFile("GMC Encountered Errors -- Exit Code :  " + cmdGMCProcess.ExitCode.ToString(), true);
                    return false;
                }

                //////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////// FIXING PRODUCTION PDF FILE VARIABLE TO BE USED BY OTHER METHODS
                //////////////////////////////////////////////////////////////////////////////////////////
                strProductionPDFFile = strProductionPDFFile.Replace(" %j", "");

                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////// IF TESTING, MOVING ALL PDFS INTO A CENTRAL DIRECTORY FOR PROOFING
                //////////////////////////////////////////////////////////////////////////////////////////
                if (bolRunInTestMode)
                {
                    try
                    {
                        if (!strOutputType.Equals("EMAILPDF"))
                        {
                            foreach (FileInfo fileProductionPDFs in new DirectoryInfo(strProductionJobFolder).GetFiles("*.pdf"))
                            {
                                fileProductionPDFs.CopyTo(strPDFProofRepository + Path.GetFileName(fileProductionPDFs.ToString()));
                            }
                        }
                    }
                    catch
                    {
                        LogFile("Unable to move production PDFs to '" + strPDFProofRepository + "'.", false);
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
                streamGMCJobBatchFile.Dispose();
                cmdGMCProcess.Close();
                cmdGMCProcess.Dispose();
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

                        strEmailSubject = "Storefront Automation - ERROR HAS OCCURED";
                        strEmailBody = "<body style=\"font-family: Arial; font-size: 12px\">" +
                                       "Please review attached log file for details." +
                                       "<br/><br/>" +
                                       "Storefront ID Number :    " + strStorefrontOrderID +
                                       "</body>";

                        strEmailAttachment = strLogFile;

                        break;

                    case "HIGHMARK":
                        strFinalEmail = strHighmarkEmailTo;
                        strFinalEmailCC = strEmailCC;

                        strEmailSubject = strProjectName + " - Region-Zip Mismatches Removed";

                        if (bolRegionMismatchThresholdReached)
                        {
                            strEmailBody = "<body style=\"font-family: Arial; font-size: 12px\">" +
                                           "Region-Zip mismatches were identified in the following job:<br/>" +
                                           strProjectName +
                                           "<br/><br/>" +
                                           "Please see attached file of records that were removed." +
                                           "<br/><br/>" +
                                           "NOTE: The mismatch threshold of " + (100 - iZipToRegionMismatchThreshold).ToString() + "% was reached/exceeded.<br/>" +
                                           "The processing of this job has been CANCELLED." +
                                           "<br/><br/>" +
                                           "Thanks!<br/>" +
                                           "Heeter" +
                                           "</body>";
                        }
                        else
                        {
                            strEmailBody = "<body style=\"font-family: Arial; font-size: 12px\">" +
                                           "Region-Zip mismatches were identified in the following job:<br/>" +
                                           strProjectName +
                                           "<br/><br/>" +
                                           "Please see attached file of records that were removed." +
                                           "<br/><br/>" +
                                           "Processing will continue with the remaining verified records." +
                                           "<br/><br/>" +
                                           "Thanks!<br/>" +
                                           "Heeter" +
                                           "</body>";
                        }

                        strEmailAttachment = strRegionErrorsRemoved;

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


        #region "    Database Log Table...    "

        static bool DatabaseLogTable()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strSQLCommand = "INSERT INTO " + strSQLLogTable + " VALUES (" +
                                   "@Storefront_Order_ID, " +
                                   "@Job_Number, " +
                                   "@Client_Code, " +
                                   "@Client_Name, " + 
                                   "@User_Name, " + 
                                   "@User_Email, " + 
                                   "@Delivery_Email, " + 
                                   "@Internal_Email_CC, " +
                                   "@Project_Name, " + 
                                   "@Output_Type, " +
                                   "@Mail_Date, " +
                                   "@Delivery_Date, " +
                                   "@Ship_Name, " + 
                                   "@Ship_Care_Of, " +
                                   "@Ship_Add_1, " + 
                                   "@Ship_Add_2, " +
                                   "@Ship_City, " +
                                   "@Ship_State, " +
                                   "@Ship_Zip, " +
                                   "@Region_Code, " +
                                   "@Input_Qty, " +
                                   "@Mailing_Qty, " +
                                   "@Input_XML_Name, " +
                                   "@Input_Data_Name, " +
                                   "@Item_Data_XML_Name, " +
                                   "@Data_Split_Value, " +
                                   "@Pass_To_Apogee, " +
                                   "@Apogee_Client_Folder, " +
                                   "@Start_Time, " +
                                   "@End_Time, " +
                                   "@Direct_Response_Order_Number, " +
                                   "@Direct_Response_CSV, " + 
                                   "@Sorted_Data, " +
                                   "@Mailing_Print_Output, " +
                                   "@Production_PDF_File, " +
                                   "@Watermark_PDF_File, " + 
                                   "@Additional_Notes)";

            SqlConnection sqlDBConnection = new SqlConnection(strSQLConnection);
            SqlCommand sqlDBCommand = new SqlCommand(strSQLCommand, sqlDBConnection);
            
            try
            {
                //////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////// CONNECTING TO SQL TABLE, AND INSERTING NEW PROCESSING INFORMAITON
                //////////////////////////////////////////////////////////////////////////////////////////
                sqlDBConnection.Open();
                
                // Creating the parameters that will be passed into the table.
                sqlDBCommand.Parameters.AddWithValue("@Storefront_Order_ID", strStorefrontOrderID);
                sqlDBCommand.Parameters.AddWithValue("@Job_Number", strJobNumber);
                sqlDBCommand.Parameters.AddWithValue("@Client_Code", strClientCode);
                sqlDBCommand.Parameters.AddWithValue("@Client_Name", strCompanyName);
                sqlDBCommand.Parameters.AddWithValue("@User_Name", strOrderFirstName + " " + strOrderLastName);
                sqlDBCommand.Parameters.AddWithValue("@User_Email", strUserEmail);
                sqlDBCommand.Parameters.AddWithValue("@Delivery_Email", strDeliveryEmail);
                sqlDBCommand.Parameters.AddWithValue("@Internal_Email_CC", strInternalEmailCC);
                sqlDBCommand.Parameters.AddWithValue("@Project_Name", strProjectName);
                sqlDBCommand.Parameters.AddWithValue("@Output_Type", strOutputType);
                sqlDBCommand.Parameters.AddWithValue("@Mail_Date", string.IsNullOrEmpty(strMailDate) ? "N/A" : strMailDate);
                sqlDBCommand.Parameters.AddWithValue("@Delivery_Date", string.IsNullOrEmpty(strDeliveryDate) ? "N/A" : strDeliveryDate);
                sqlDBCommand.Parameters.AddWithValue("@Ship_Name", string.IsNullOrEmpty(strShipName) ? "N/A" : strShipName);
                sqlDBCommand.Parameters.AddWithValue("@Ship_Care_Of", string.IsNullOrEmpty(strShipCareOf) ? "N/A" : strShipCareOf);
                sqlDBCommand.Parameters.AddWithValue("@Ship_Add_1", string.IsNullOrEmpty(strShipAdd1) ? "N/A" : strShipAdd1);
                sqlDBCommand.Parameters.AddWithValue("@Ship_Add_2", string.IsNullOrEmpty(strShipAdd2) ? "N/A" : strShipAdd2);
                sqlDBCommand.Parameters.AddWithValue("@Ship_City", string.IsNullOrEmpty(strShipCity) ? "N/A" : strShipCity);
                sqlDBCommand.Parameters.AddWithValue("@Ship_State", string.IsNullOrEmpty(strShipState) ? "N/A" : strShipState);
                sqlDBCommand.Parameters.AddWithValue("@Ship_Zip", string.IsNullOrEmpty(strShipZip) ? "N/A" : strShipZip);
                sqlDBCommand.Parameters.AddWithValue("@Region_Code", string.IsNullOrEmpty(strRegionCode) ? "N/A" : strRegionCode);
                sqlDBCommand.Parameters.AddWithValue("@Input_Qty", iInputQuantity);
                sqlDBCommand.Parameters.AddWithValue("@Mailing_Qty", iMailingRecords);
                sqlDBCommand.Parameters.AddWithValue("@Input_XML_Name", Path.GetFileName(strInputXMLName));
                sqlDBCommand.Parameters.AddWithValue("@Input_Data_Name", string.IsNullOrEmpty(Path.GetFileName(strInputDataName)) ? "N/A" : Path.GetFileName(strInputDataName));
                sqlDBCommand.Parameters.AddWithValue("@Item_Data_XML_Name", strItemDataXMLNameOnly);
                sqlDBCommand.Parameters.AddWithValue("@Data_Split_Value", iDataSplitValue);
                sqlDBCommand.Parameters.AddWithValue("@Pass_To_Apogee", bolPassToApogee.Equals(true) ? 'Y' : 'N');
                sqlDBCommand.Parameters.AddWithValue("@Apogee_Client_Folder", bolPassToApogee.Equals(true) ? strApogeeClientFolder : "N/A");
                sqlDBCommand.Parameters.AddWithValue("@Start_Time", dtStartTime.ToString());
                sqlDBCommand.Parameters.AddWithValue("@End_Time", dtEndTime.ToString());
                sqlDBCommand.Parameters.AddWithValue("@Direct_Response_Order_Number", strDirectResponseOrderNumber);
                sqlDBCommand.Parameters.AddWithValue("@Direct_Response_CSV", strDirectResponseCSV);
                sqlDBCommand.Parameters.AddWithValue("@Sorted_Data", string.IsNullOrEmpty(strSortedData) ? "N/A" : strSortedData);
                sqlDBCommand.Parameters.AddWithValue("@Mailing_Print_Output", string.IsNullOrEmpty(strMailingPrintOutput) ? "N/A" : strMailingPrintOutput);
                sqlDBCommand.Parameters.AddWithValue("@Production_PDF_File", string.IsNullOrEmpty(strProductionPDFFile) ? "N/A" : strProductionPDFFile);
                sqlDBCommand.Parameters.AddWithValue("@Watermark_PDF_File", string.IsNullOrEmpty(strWatermarkPDFFile) ? "N/A" : strWatermarkPDFFile);
                sqlDBCommand.Parameters.AddWithValue("@Additional_Notes", strAdditionalNotes);

                // Executing the SQL query.
                sqlDBCommand.ExecuteNonQuery();
            }

            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            finally
            {
                sqlDBCommand.Dispose();
                sqlDBConnection.Close();
                sqlDBConnection.Dispose();
            }

            return true;
        }

        #endregion


        #region "    Database Post Process Table...    "

        static bool DatabasePostProcessTable()
        {
            //////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////// INITIALIZING VARIABLES
            //////////////////////////////////////////////////////////////////////////////////////////
            string strSQLCommand = "INSERT INTO " + strSQLPostProcessTable + " VALUES (" +
                                   "@Direct_Response_Order_Number, " +
                                   "@Project_Name, " +
                                   "@Quantity, " +
                                   "@Product_ID, " +
                                   "@Product_Price, " +
                                   "@Postage_Amount, " +
                                   "@Shipping_Amount, " +
                                   "@Tax_Amount, " +
                                   "@Order_Fee, " +
                                   "@Credit_Card_Required, " +
                                   "@Job_Type, " +
                                   "@Job_Status, " +
                                   "@Email_Sent)";

            SqlConnection sqlDBConnection = new SqlConnection(strSQLConnection);
            SqlCommand sqlDBCommand = new SqlCommand(strSQLCommand, sqlDBConnection);

            //////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////// CONNECTING TO SQL TABLE, AND INSERTING NEW PROCESSING INFORMATION
            //////////////////////////////////////////////////////////////////////////////////////////
            try
            {
                sqlDBConnection.Open();

                // Creating the parameters that will be passed into the table.
                sqlDBCommand.Parameters.AddWithValue("@Direct_Response_Order_Number", strDirectResponseOrderNumber);
                sqlDBCommand.Parameters.AddWithValue("@Project_Name", strProjectName);
                sqlDBCommand.Parameters.AddWithValue("@Quantity", (strOutputType.Equals("PRINTMAIL") ? iMailingRecords : iInputQuantity));
                sqlDBCommand.Parameters.AddWithValue("@Product_ID", strProductID);
                sqlDBCommand.Parameters.AddWithValue("@Product_Price", (strOutputType.Equals("EMAILPDF") ? 0.0 : dblProductPrice));
                sqlDBCommand.Parameters.AddWithValue("@Postage_Amount", dblPostageAmount);
                sqlDBCommand.Parameters.AddWithValue("@Shipping_Amount", dblShippingAmount);
                sqlDBCommand.Parameters.AddWithValue("@Tax_Amount", dblTaxAmount);
                sqlDBCommand.Parameters.AddWithValue("@Order_Fee", dblOrderFee);
                sqlDBCommand.Parameters.AddWithValue("@Credit_Card_Required", strCreditCardRequired);
                sqlDBCommand.Parameters.AddWithValue("@Job_Type", strOutputType);
                sqlDBCommand.Parameters.AddWithValue("@Job_Status", strJobStatus);
                sqlDBCommand.Parameters.AddWithValue("@Email_Sent", strJobStatus.Equals("PAYMENT NEEDED") ? "N" : "N/A");

                // Executing the SQL query.
                sqlDBCommand.ExecuteNonQuery();
            }

            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            finally
            {
                sqlDBCommand.Dispose();
                sqlDBConnection.Close();
                sqlDBConnection.Dispose();
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


        #region "    Copy File Using Credentials...    "

        public static bool CopyFileUsingCredentials(string strFileToCopyFrom, string strFileToCopyTo, string strUsername, string strPassword)
        {
            try
            {
                string strCopyToDirectory = Path.GetDirectoryName(strFileToCopyTo).Trim();
                string strCopyToFilename = Path.GetFileName(strFileToCopyTo);

                if (!strCopyToDirectory.EndsWith("\\"))
                    strCopyToFilename = "\\" + strCopyToFilename;

                string command = "NET USE " + strCopyToDirectory + " /delete";
                ExecuteCommand(command, 5000);

                command = "NET USE " + strCopyToDirectory + " /user:" + strUsername + " " + strPassword;
                ExecuteCommand(command, 5000);

                command = " copy \"" + strFileToCopyFrom + "\"  \"" + strCopyToDirectory + strCopyToFilename + "\"";
                ExecuteCommand(command, 5000);

                command = "NET USE " + strCopyToDirectory + " /delete";
                ExecuteCommand(command, 5000);
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion


        #region "    Execute Command...    "

        public static bool ExecuteCommand(string strCommand, int iTimeout)
        {
            try
            {
                ProcessStartInfo processRunStartInfo = new ProcessStartInfo("cmd.exe", "/C " + strCommand)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WorkingDirectory = "C:\\",
                };

                Process processRunCommand = Process.Start(processRunStartInfo);
                processRunCommand.WaitForExit(iTimeout);
                processRunCommand.Close();
            }
            catch (Exception exception)
            {
                LogFile(exception.ToString(), true);
                return false;
            }

            return true;
        }

        #endregion
    }
}