using Newtonsoft.Json;
using System.Data;
using System.Text;
using System.Web;
using RestSharp;

namespace Dynamics_Oracle_UnitBilling
{
    class Program
    {
        static string AppName = "OracleUnitBillingAutomation";

        static string Oracle_Url = System.Configuration.ConfigurationManager.AppSettings["Oracle_Url"]!.ToString();
        static string UnitBilling_Sub_Url = System.Configuration.ConfigurationManager.AppSettings["UnitBilling_Sub_Url"]!.ToString();
        static string PPM_SubUrl = System.Configuration.ConfigurationManager.AppSettings["PPM_Sub_Url"]!.ToString();
        static string PPM_EssJob_Sub_Url = System.Configuration.ConfigurationManager.AppSettings["PPM_EssJob_Sub_Url"]!.ToString();

        static string UserName = System.Configuration.ConfigurationManager.AppSettings["ServiceUserId"]!.ToString();
        static string Password = clsTools.Decrypt(System.Configuration.ConfigurationManager.AppSettings["ServicePassword"]!.ToString(), true);
        static string ServiceTimeout = System.Configuration.ConfigurationManager.AppSettings["ServiceTimeout"]!.ToString();
        
        static readonly string Dynamics_Url = System.Configuration.ConfigurationManager.AppSettings["Dynamics_Url"]!.ToString();
        static readonly int ExceptionOnScreen = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["ExceptionOnScreen"]!.ToString());

        #region Main
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("DynamicsPikeService - " + AppName + " - Started");

                CreateUnitBillingEvents();

                UpdateUnitBillingEvents();

                Console.WriteLine("DynamicsPikeService - " + AppName + " - Completed");
            }
            catch (Exception exp)
            {
                Console.WriteLine(exp.Message.ToString());
            }
        }
        #endregion

        #region CreateUnitBillingEvents
        public static void CreateUnitBillingEvents()
        {
            // File name  
            FileStream? stream = null;            
            string? filePath = System.Configuration.ConfigurationManager.AppSettings["FilePath"]!.ToString();
            string? FolderName = System.Configuration.ConfigurationManager.AppSettings["FolderName"]!.ToString();
            System.IO.Directory.CreateDirectory(filePath + FolderName);

            string strPayload = "";
            string strHPayload = "";
            string strHeaderEntity = "hsl_oracleintegrationbatch_headers";
            string strEntity = "hsl_oracleintegrationbatch_detailedrecords";
            string hsl_oracle_batch_status_header = "";
            string hsl_oracle_status_details = "";
            string hsl_oracle_message_details = "";
            string hsl_ppm_pc_transaction_id_details = "";
            string hsl_unprocessedtransactionreferenceid_details = "";

            string DetailRecordId = "";
            string IntegrationBatchHeaderId = "";
            string ExpenditureBatchName = "";            

            DataSet dsBatchHeader = clsDAL.Dynamics_UnitBilling_BatchHeader_Create();

            if (dsBatchHeader.Tables[0].Rows.Count > 0 )
            {
                string? fileName = filePath + FolderName + "\\UnitBilling_Log_Create_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                stream = new FileStream(fileName, FileMode.OpenOrCreate);
                // Create a StreamWriter from FileStream  
                StreamWriter writer = new StreamWriter(stream, Encoding.UTF8);
                writer.AutoFlush = true;

                writer.WriteLine("DynamicsPikeService (Create) - " + AppName + " - Started");
                

                for (int i = 0; i < dsBatchHeader.Tables[0].Rows.Count;i++) 
                {
                    int HeaderException_Flag = 0;
                    ExpenditureBatchName = dsBatchHeader.Tables[0].Rows[i]["ExpenditureBatch"].ToString()!;
                    PPM_EXEC_ESSJOB(ExpenditureBatchName);
                    IntegrationBatchHeaderId = dsBatchHeader.Tables[0].Rows[i]["IntegrationBatchHeaderId"].ToString()!;

                    writer.WriteLine("Featching Header Records for Batch = " + ExpenditureBatchName);
                    writer.WriteLine("Processing Header Count " + (i + 1).ToString() + " out of " + dsBatchHeader.Tables[0].Rows.Count.ToString());
                    Console.WriteLine("Processing Header Count " + (i + 1).ToString() + " out of " + dsBatchHeader.Tables[0].Rows.Count.ToString());
                    writer.WriteLine("Updating the Header Record Status to In Progress for "+ hsl_oracle_batch_status_header);

                    //Updating the Header Table Status to In Process
                    hsl_oracle_batch_status_header = "In Process";
                    strHPayload = "{ "
                                     + "\"hsl_oracle_batch_status\":\"" + hsl_oracle_batch_status_header + "\"}";

                    UpdateDynamics(strHeaderEntity, IntegrationBatchHeaderId!, strHPayload);
                    DataSet dsDetails = clsDAL.Dynamics_UniBilling_List_Create(ExpenditureBatchName);

                    try
                    {
                        if (dsDetails.Tables[0].Rows.Count > 0)
                        {
                            for (int j = 0; j < dsDetails.Tables[0].Rows.Count; j++)
                            {
                                try
                                {
                                    Console.WriteLine("Processing Detail Count " + (j + 1).ToString() + " out of " + dsDetails.Tables[0].Rows.Count.ToString());
                                    writer.WriteLine("Processing Detail Count " + (j + 1).ToString() + " out of " + dsDetails.Tables[0].Rows.Count.ToString());

                                    int RecordFlag = 1;

                                    byte[] toEncodeAsBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(UserName + ":" + Password);
                                    string credentials = System.Convert.ToBase64String(toEncodeAsBytes);

                                    DetailRecordId = dsDetails.Tables[0].Rows[j]["DetailRecordId"].ToString()!.ToLower();                                    

                                    string? OriginalTransactionReference = "Replicon_" + DetailRecordId;


                                    #region Check if the record got created
                                    writer.WriteLine("Checking if the Transaction is already pushed to PPM for Detail Record Id "+ DetailRecordId);
                                    string? PPM_QueryParam = "?q=OriginalTransactionReference=" + OriginalTransactionReference + ";NetZeroItemFlag is null&expand=ProjectStandardCostCollectionFlexFields";
                                    writer.WriteLine("Payload for Getting the Transaction Number " + PPM_QueryParam);

                                    var PPMoptions = new RestClientOptions(Oracle_Url)
                                    {
                                        MaxTimeout = -1,
                                    };
                                    var PPMclient = new RestClient(PPMoptions);
                                    var PPMrequest = new RestRequest(PPM_SubUrl + PPM_QueryParam, Method.Get);
                                    PPMrequest.AddHeader("Content-Type", "application/json");
                                    PPMrequest.AddHeader("Authorization", "Basic " + credentials);

                                    RestResponse PPMresponse = PPMclient.Execute(PPMrequest);
                                    writer.WriteLine("PPM Response to get TransactionNumber : " + PPMresponse.Content);

                                    if (PPMresponse.StatusCode == System.Net.HttpStatusCode.OK)
                                    {
                                        //Update the PPM Status
                                        string Result = PPMresponse.Content!.ToString();
                                        dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(PPMresponse.Content!)!;
                                        string TransactionNumber = "";
                                        if (dyArray.count != 0)
                                            TransactionNumber = dyArray.items[0].TransactionNumber.ToString();

                                        writer.WriteLine("TransactionNumber : " + TransactionNumber);
                                        Console.WriteLine("TransactionNumber : " + TransactionNumber);

                                        if (TransactionNumber != null && TransactionNumber != "" && TransactionNumber != "0")
                                        {
                                            RecordFlag = 0;

                                            writer.WriteLine("Updating the Details Record with the Transaction Number into Dynamics for Detail ID= " + DetailRecordId);

                                            hsl_oracle_status_details = "Success";
                                            hsl_oracle_message_details = "Transaction Number Updated";
                                            hsl_ppm_pc_transaction_id_details = TransactionNumber;

                                            //Update the PPM Status
                                            strPayload = "{ "
                                            + "\"hsl_oracle_status\":\"" + hsl_oracle_status_details + "\","
                                            + "\"hsl_oracle_message\":\"" + hsl_oracle_message_details + "\","
                                            + "\"hsl_ppm_pc_transaction_id\":\"" + hsl_ppm_pc_transaction_id_details + "\"}";

                                            UpdateDynamics(strEntity, DetailRecordId!, strPayload);
                                        }
                                    }

                                    #endregion

                                    if (RecordFlag == 1)
                                    {

                                        Console.WriteLine("Processing Unit Billing Detail Events Records - " + j.ToString() + " - out of " + dsDetails.Tables[0].Rows.Count.ToString());
                                        writer.WriteLine("Processing Unit Billing Detail Events Records - " + j.ToString() + " - out of " + dsDetails.Tables[0].Rows.Count.ToString());

                                        
                                        string? BusinessUnit = dsDetails.Tables[0].Rows[j]["BusinessUnit"].ToString();
                                        string? TransactionSource = dsDetails.Tables[0].Rows[j]["TransactionSource"].ToString();

                                        string? Document = dsDetails.Tables[0].Rows[j]["Document"].ToString();
                                        string? DocumentEntry = dsDetails.Tables[0].Rows[j]["DocumentEntry"].ToString();


                                        string? NonlaborResourceID = dsDetails.Tables[0].Rows[j]["NonlaborResourceID"].ToString();
                                        string? NonlaborResource = dsDetails.Tables[0].Rows[j]["NonlaborResource"].ToString();
                                        string? NonlaborResourceOrganization = dsDetails.Tables[0].Rows[j]["NonlaborResourceOrganization"].ToString();
                                        string? TransactionCurrencyCode = dsDetails.Tables[0].Rows[j]["TransactionCurrencyCode"].ToString();
                                        
                                        string? Quantity = dsDetails.Tables[0].Rows[j]["Quantity"].ToString();
                                        //string? UnitOfMeasure = dsDetails.Tables[0].Rows[j]["UnitOfMeasure"].ToString();
                                        string? UnitOfMeasure = "Hours";

                                        DateTime dtExpenditureDate = Convert.ToDateTime(dsDetails.Tables[0].Rows[j]["EXPENDITURE_ITEM_DATE"].ToString());
                                        string? ExpenditureDate = dtExpenditureDate.ToString("yyyy-MM-dd");

                                        string? ProjectId = dsDetails.Tables[0].Rows[j]["PROJECT_ID"].ToString();
                                        string? TaskId = dsDetails.Tables[0].Rows[j]["TASK_ID"].ToString();
                                        string? ExpenditureTypeId = dsDetails.Tables[0].Rows[j]["EXPENDITURE_TYPE_ID_Display"].ToString();
                                        string? OrganizationId = dsDetails.Tables[0].Rows[j]["ORGANIZATION_ID_Display"].ToString();


                                        writer.WriteLine("Updating the Detail Record Status to In Progress for " + DetailRecordId);
                                        strPayload = "{\"hsl_oracle_status\":\"In Process\" }";
                                        UpdateDynamics(strEntity, DetailRecordId!, strPayload);


                                        //PPM UPC CREATE
                                        var options = new RestClientOptions(Oracle_Url)
                                        {
                                            MaxTimeout = -1,
                                        };
                                        var client = new RestClient(options);
                                        var request = new RestRequest(UnitBilling_Sub_Url, Method.Post);
                                        request.AddHeader("Content-Type", "application/json");
                                        request.AddHeader("Authorization", "Basic " + credentials);

                                        var body = "{"

                                                        + "\"ExpenditureBatch\":\"" + ExpenditureBatchName + "\","
                                                        + "\"BusinessUnit\":\"" + BusinessUnit + "\","
                                                        + "\"TransactionSource\":\"" + TransactionSource + "\","
                                                        + "\"Document\":\"" + Document + "\","
                                                        + "\"DocumentEntry\":\"" + DocumentEntry + "\","
                                                        + "\"NonlaborResourceId\":\"" + NonlaborResourceID + "\","
                                                        + "\"NonlaborResource\":\"" + NonlaborResource + "\","
                                                        + "\"NonlaborResourceOrganization\":\"" + NonlaborResourceOrganization + "\","
                                                        + "\"TransactionCurrencyCode\":\"" + TransactionCurrencyCode + "\","
                                                        + "\"OriginalTransactionReference\":\"" + OriginalTransactionReference + "\","
                                                        + "\"Quantity\":\"" + Quantity + "\","
                                                        + "\"UnitOfMeasure\":\"" + UnitOfMeasure + "\","
                                                        + "\"ProjectStandardCostCollectionFlexfields\": ["
                                                                + "{"
                                                                + "\"_EXPENDITURE_ITEM_DATE\":\"" + ExpenditureDate + "\","
                                                            + "\"_PROJECT_ID\":\"" + ProjectId + "\","
                                                            + "\"_TASK_ID\":\"" + TaskId + "\","
                                                            + "\"_EXPENDITURE_TYPE_ID_Display\":\"" + ExpenditureTypeId + "\","
                                                            + "\"_ORGANIZATION_ID_Display\":\"" + OrganizationId + "\""
                                                            + "}]"
                                                            + "}";



                                        request.AddStringBody(body, DataFormat.Json);
                                        RestResponse response = client.Execute(request);

                                        writer.WriteLine("Payload for Detail Transaction Record :  " + body);

                                        if (response.Content != "" && response.StatusCode.ToString() == "Created")
                                        {
                                            dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!)!;
                                            writer.WriteLine("Oracle Response for Detail Transaction Record :  " + response.Content!);

                                            writer.WriteLine("Updating the Detail Record Status in Dynamics along with other values for " + DetailRecordId);

                                            hsl_oracle_status_details = "Success";
                                            hsl_oracle_message_details = "Sent to UPC";
                                            hsl_unprocessedtransactionreferenceid_details = dyArray.UnprocessedTransactionReferenceId.ToString(); 

                                            //Update the PPM Status
                                            strPayload = "{ "
                                           + "\"hsl_oracle_status\":\"" + hsl_oracle_status_details + "\","
                                           + "\"hsl_oracle_message\":\"" + hsl_oracle_message_details + "\","
                                           + "\"hsl_unprocessedtransactionreferenceid\":\"" + hsl_unprocessedtransactionreferenceid_details + "\"}";

                                            writer.WriteLine("Updating the Dynamics Detail Record Columns with the attached payload " + strPayload);
                                            UpdateDynamics(strEntity, DetailRecordId!, strPayload);
                                        }
                                    }
                                }
                                catch (Exception exp)
                                {
                                    HeaderException_Flag = 1;

                                    writer.WriteLine("Oralce Response was Errored out with the following Exceptions "+ exp.Message);
                                    Console.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);

                                    hsl_oracle_status_details = "Error";
                                    hsl_oracle_message_details = exp.Message.ToString();

                                    //Update the PPM Status
                                    strPayload = "{ "
                                   + "\"hsl_oracle_status\":\"" + hsl_oracle_status_details + "\","
                                   + "\"hsl_oracle_message\":\"" + hsl_oracle_message_details + "\"}";

                                    writer.WriteLine("Updating the Dynamics Detail Record Columns with the attached payload (Exception) " + strPayload);
                                    UpdateDynamics(strEntity, DetailRecordId!, strPayload);
                                }
                            }
                        }

                    }
                    catch (Exception exp)
                    {
                        HeaderException_Flag = 1;
                        writer.WriteLine("Application Exception " + exp.Message);
                        Console.WriteLine("Application Exception " + exp.Message);
                    }

                    if (HeaderException_Flag == 0)
                    {
                        hsl_oracle_batch_status_header = "Sent to UPC";
                        strHPayload = "{ "
                                     + "\"hsl_oracle_batch_status\":\"" + hsl_oracle_batch_status_header + "\"}";
                        UpdateDynamics(strHeaderEntity, IntegrationBatchHeaderId!, strHPayload);
                        writer.WriteLine("Dynamics Success Update for Header Table Payload :  " + strHPayload);
                    }
                    PPM_EXEC_ESSJOB(ExpenditureBatchName);
                }

                writer.WriteLine("DynamicsPikeService (Create) - " + AppName + " - Completed");
            }
        }
        #endregion


        #region UpdateUnitBillingEvents
        public static void UpdateUnitBillingEvents()
        {
            Console.WriteLine("DynamicsPikeService (Update) - " + AppName + " - Started");

            FileStream? stream = null;
            // File name  
            string? filePath = System.Configuration.ConfigurationManager.AppSettings["FilePath"]!.ToString();
            string? FolderName = System.Configuration.ConfigurationManager.AppSettings["FolderName"]!.ToString();
            System.IO.Directory.CreateDirectory(filePath + FolderName);

            string hsl_oracle_batch_status_header = "";
            string hsl_oracle_status_details = "";
            string hsl_oracle_message_details = "";
            string hsl_ppm_pc_transaction_id_details = "";
            string strPayload = "";
            string strHPayload = "";
            string strHeaderEntity = "hsl_oracleintegrationbatch_headers";
            string strEntity = "hsl_oracleintegrationbatch_detailedrecords";

            string DetailRecordId = "";
            string IntegrationBatchHeaderId = "";
            string ExpenditureBatchName = "";

            DataSet dsBatchHeader = clsDAL.Dynamics_UnitBilling_BatchHeader_Update();

            if (dsBatchHeader.Tables[0].Rows.Count > 0)
            {
                string? fileName = filePath + FolderName + "\\UnitBilling_Log_Update_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                stream = new FileStream(fileName, FileMode.OpenOrCreate);
                // Create a StreamWriter from FileStream  
                StreamWriter writer = new StreamWriter(stream, Encoding.UTF8);
                writer.AutoFlush = true;

                writer.WriteLine("DynamicsPikeService (Update) - " + AppName + " - Started");


                for (int i = 0; i < dsBatchHeader.Tables[0].Rows.Count; i++)
                {
                    int HeaderException_Flag = 0;
                    ExpenditureBatchName = dsBatchHeader.Tables[0].Rows[i]["ExpenditureBatch"].ToString()!;
                    PPM_EXEC_ESSJOB(ExpenditureBatchName);
                    IntegrationBatchHeaderId = dsBatchHeader.Tables[0].Rows[i]["IntegrationBatchHeaderId"].ToString()!;

                    writer.WriteLine("Featching Header Records for Batch = " + ExpenditureBatchName);
                    writer.WriteLine("Processing Header Count " + (i + 1).ToString() + " out of " + dsBatchHeader.Tables[0].Rows.Count.ToString());
                    Console.WriteLine("Processing Header Count " + (i + 1).ToString() + " out of " + dsBatchHeader.Tables[0].Rows.Count.ToString());

                    DataSet dsDetails = clsDAL.Dynamics_UniBilling_List_Update(ExpenditureBatchName);
                    try
                    {
                        writer.WriteLine("DynamicsPikeService - " + AppName + " - Total Unit Billing Events " + dsBatchHeader.Tables[0].Rows.Count.ToString());
                        Console.WriteLine("Processing Header Count " + (i + 1).ToString() + " out of " + dsBatchHeader.Tables[0].Rows.Count.ToString());
                        Console.WriteLine("---------------------------------------");


                        if (dsDetails.Tables[0].Rows.Count > 0)
                        {
                            for (int j = 0; j < dsDetails.Tables[0].Rows.Count; j++)
                            {
                                try
                                {
                                    Console.WriteLine("Processing Detail Count " + (j + 1).ToString() + " out of " + dsDetails.Tables[0].Rows.Count.ToString());
                                    writer.WriteLine("Processing Detail Count " + (j + 1).ToString() + " out of " + dsDetails.Tables[0].Rows.Count.ToString());

                                    byte[] toEncodeAsBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(UserName + ":" + Password);
                                    string credentials = System.Convert.ToBase64String(toEncodeAsBytes);

                                    DetailRecordId = dsDetails.Tables[0].Rows[j]["DetailRecordId"].ToString()!.ToLower();

                                    string? OriginalTransactionReference = "Replicon_" + DetailRecordId;


                                    writer.WriteLine("Checking if the Transaction is already pushed to PPM for Detail Record Id " + DetailRecordId);
                                    #region Check if the record got created
                                    string? PPM_QueryParam = "?q=OriginalTransactionReference=" + OriginalTransactionReference + ";NetZeroItemFlag is null&expand=ProjectStandardCostCollectionFlexFields";

                                    writer.WriteLine("Payload for Getting the Transaction Number " + PPM_QueryParam);

                                    var PPMoptions = new RestClientOptions(Oracle_Url)
                                    {
                                        MaxTimeout = -1,
                                    };
                                    var PPMclient = new RestClient(PPMoptions);
                                    var PPMrequest = new RestRequest(PPM_SubUrl + PPM_QueryParam, Method.Get);
                                    PPMrequest.AddHeader("Content-Type", "application/json");
                                    PPMrequest.AddHeader("Authorization", "Basic " + credentials);

                                    RestResponse PPMresponse = PPMclient.Execute(PPMrequest);
                                    writer.WriteLine("PPM Response to get TransactionNumber : " + PPMresponse.Content);

                                    //Console.WriteLine(PPMresponse.Content);
                                    if (PPMresponse.StatusCode == System.Net.HttpStatusCode.OK)
                                    {

                                        //Update the PPM Status
                                        string Result = PPMresponse.Content!.ToString();
                                        dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(PPMresponse.Content!)!;
                                        string TransactionNumber = "";
                                        if (dyArray.count != 0)
                                            TransactionNumber = dyArray.items[0].TransactionNumber.ToString();

                                        writer.WriteLine("TransactionNumber : " + TransactionNumber);
                                        Console.WriteLine("TransactionNumber : " + TransactionNumber);
                                        Console.WriteLine("");

                                        if (TransactionNumber != null && TransactionNumber != "" && TransactionNumber != "0")
                                        {
                                            writer.WriteLine("Updating the Details Record with the Transaction Number into Dynamics for Detail ID= " + DetailRecordId);

                                            // Prepare web request and pass token.
                                            //Update the PPM Status details

                                            hsl_oracle_status_details = "Success";
                                            hsl_oracle_message_details = "Transaction Number Updated";
                                            hsl_ppm_pc_transaction_id_details = TransactionNumber;

                                            //Update the PPM Status
                                            strPayload = "{ "
                                           + "\"hsl_oracle_status\":\"" + hsl_oracle_status_details + "\","
                                           + "\"hsl_oracle_message\":\"" + hsl_oracle_message_details + "\","
                                           + "\"hsl_ppm_pc_transaction_id\":\"" + hsl_ppm_pc_transaction_id_details + "\"}";

                                            UpdateDynamics(strEntity, DetailRecordId!, strPayload);

                                        }
                                    }

                                    #endregion


                                }
                                catch (Exception exp)
                                {
                                    HeaderException_Flag = 1;

                                    writer.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);
                                    Console.WriteLine("Oralce Response was Errored out with the following Exceptions " + exp.Message);

                                    hsl_oracle_status_details = "Error";
                                    hsl_oracle_message_details = exp.Message.ToString();

                                    //Update the PPM Status
                                    strPayload = "{ "
                                   + "\"hsl_oracle_status\":\"" + hsl_oracle_status_details + "\","
                                   + "\"hsl_oracle_message\":\"" + hsl_oracle_message_details + "\"}";

                                    writer.WriteLine("Updating the Dynamics Detail Record Columns with the attached payload (Exception) " + strPayload);
                                    UpdateDynamics(strEntity, DetailRecordId!, strPayload);
                                }
                            }
                        }

                    }
                    catch (Exception exp)
                    {
                        HeaderException_Flag = 1;
                        writer.WriteLine("Application Exception " + exp.Message);
                        Console.WriteLine("Application Exception " + exp.Message);
                    }

                    if (HeaderException_Flag == 0)
                    {
                        
                        DataSet dsHeaderCount = clsDAL.Dynamics_HeaderCount(ExpenditureBatchName);

                        if (dsHeaderCount.Tables[0].Rows[0]["RecordCount"].ToString() == "0")
                        {
                            hsl_oracle_batch_status_header = "Success";

                            strHPayload = "{ "
                                         + "\"hsl_oracle_batch_status\":\"" + hsl_oracle_batch_status_header + "\"}";
                            UpdateDynamics(strHeaderEntity, IntegrationBatchHeaderId!, strHPayload);
                            writer.WriteLine("Dynamics Success Update for Header Table Payload :  " + strHPayload);
                        }
                    }
                }

            }
        }
        #endregion


        #region PPM_EXEC_ESSJOB
        public static void PPM_EXEC_ESSJOB(string ExpenditureBatch)
        {

            string? PPM_EssJob_Url = System.Configuration.ConfigurationManager.AppSettings["Oracle_Url"]!.ToString();
            string? PPM_EssJob_Sub_Url = System.Configuration.ConfigurationManager.AppSettings["PPM_EssJob_Sub_Url"]!.ToString();

            Console.WriteLine("DynamicsPikeService - PPM_ESSJOB Started :" + DateTime.Now.ToString("yyyy-MM-ddTHH\\_mm"));

            string? strBU_Name = System.Configuration.ConfigurationManager.AppSettings["Business_Unit"]!.ToString();
            string? strEss_Units = System.Configuration.ConfigurationManager.AppSettings["ESS_Job_Units"]!.ToString();
            string? strCost_Units = System.Configuration.ConfigurationManager.AppSettings["ESS_Cost_Units"]!.ToString();
            string? ESSParameters_Batch = "Pike Business Unit," + strBU_Name + ",IMPORT_AND_PROCESS,PREV_NOT_IMPORTED,," + strEss_Units + "," + strCost_Units + "," + ExpenditureBatch + ",,,,,ORA_PJC_DETAIL";
            Console.WriteLine("ESSParameters_Units : " + ESSParameters_Batch);
            string strESS_Job_Units = PPM_ESSJOB(PPM_EssJob_Url, PPM_EssJob_Sub_Url, UserName, Password, ESSParameters_Batch);
            Console.WriteLine("PPM ESS JOB Created Successful!!");
            Console.WriteLine("");

        }
        #endregion

        #region PPM_ESSJOB
        public static string PPM_ESSJOB(string MainUrl, string SubUrl, string username, string pwd, string ESSParameters)
        {
            string? UserJSONString = string.Empty;

            string? TransactionType = "";
            int? Transaction_Id = 0;

            string? OperationName = "submitESSJobRequest";
            string? JobPackageName = "oracle/apps/ess/projects/costing/transactions/onestop";
            string? JobDefName = "ImportProcessParallelEssJob";
            //string ESSParameters = "Pike Engineering Business Unit,300001588285232,IMPORT_AND_PROCESS,PREV_NOT_IMPORTED,,300000007509139,," + Expenditure_Batch + ",,,,,ORA_PJC_DETAIL";
            string? ReqstId = "null";
            string? strEss_Job_Result = "";
            UserJSONString = "{\"OperationName\":\"" + OperationName + "\", \"JobPackageName\":\"" + JobPackageName + "\", \"JobDefName\": \"" + JobDefName + "\",\"ESSParameters\":\"" + ESSParameters + "\",\"ReqstId\":\"" + ReqstId + "\" }";
            Console.WriteLine("Oracle Request : " + UserJSONString);

            var options = new RestClientOptions(MainUrl)
            {
                MaxTimeout = -1,
            };
            var client = new RestClient(options);
            var request = new RestRequest(SubUrl, Method.Post);

            byte[] toEncodeAsBytes1 = System.Text.ASCIIEncoding.ASCII.GetBytes(username + ":" + pwd);
            string? credentials1 = System.Convert.ToBase64String(toEncodeAsBytes1);
            request.AddHeader("Authorization", "Basic " + credentials1);

            // request.AddHeader("Content-Type", "application/json");
            //request.AddParameter("application/json", UserJSONString, ParameterType.RequestBody);
            request.AddHeader("Content-Type", "application/vnd.oracle.adf.resourceitem+json");
            request.AddStringBody(UserJSONString, DataFormat.Json);
            RestResponse response = client.Execute(request);

            if (response.Content != "" && response.StatusCode.ToString() == "Created" || response.StatusCode.ToString() == "NoContent")
            {
                dynamic dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!)!;

                TransactionType = dyArray.JobDefName!.ToString();
                Transaction_Id = dyArray.ReqstId;
                strEss_Job_Result = Transaction_Id.ToString();
                Console.WriteLine("Oracle Transaction ID :  " + strEss_Job_Result);               
            }

            return strEss_Job_Result!;
        }
        #endregion       

        #region UpdateDynamics
        public static void UpdateDynamics(string strEntity, string strGUID, string strPayload)
        {
            string AccessToken = GetToken();
            try
            {

                var options = new RestClientOptions(Dynamics_Url)
                {
                    MaxTimeout = -1,
                };
                var client = new RestClient(options);
                var request = new RestRequest("/api/data/v9.2/" + strEntity + "(" + strGUID + ")", Method.Patch);
                request.AddHeader("Authorization", "Bearer " + AccessToken);
                request.AddHeader("Accept", "application/json");
                request.AddHeader("Content-Type", "application/json");
                var body = strPayload;

                request.AddStringBody(body, DataFormat.Json);
                RestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);

            }
            catch (Exception exp)
            {
                var options = new RestClientOptions(Dynamics_Url)
                {
                    MaxTimeout = -1,
                };
                var client = new RestClient(options);
                var request = new RestRequest("/api/data/v9.2/" + strEntity + "(" + strGUID + ")", Method.Patch);
                request.AddHeader("Authorization", "Bearer " + AccessToken);
                request.AddHeader("Accept", "application/json");
                request.AddHeader("Content-Type", "application/json");
                strPayload = "{ "
                                   + "\"hsl_oracleresponse\":\"No\","
                                   + "\"hsl_oracle_status\":\"Error\","
                                   + "\"hsl_oracle_message\":\"Dynamics Service Exception\"}";
                var body = strPayload;

                request.AddStringBody(body, DataFormat.Json);
                RestResponse response = client.Execute(request);
                //Console.WriteLine(response.Content);

                Console.WriteLine(exp.Message.ToString());
                if (ExceptionOnScreen == 1)
                {
                    Console.WriteLine("Press any key to Continue");
                    Console.ReadKey();
                }
            }
        }
        #endregion

        #region Generate Token
        public static string GetToken()
        {
            string Client_Id = System.Configuration.ConfigurationManager.AppSettings["Client_Id"]!.ToString();
            string Client_Secret = clsTools.Decrypt(System.Configuration.ConfigurationManager.AppSettings["Client_Secret"]!.ToString(), true);
            string Grant_Type = System.Configuration.ConfigurationManager.AppSettings["Grant_Type"]!.ToString();
            string Realm = System.Configuration.ConfigurationManager.AppSettings["Realm"]!.ToString();
            string resourceurl = HttpUtility.UrlEncode(Dynamics_Url);
            string AccessToken = "";

            var options = new RestClientOptions("https://login.microsoftonline.com")
            {
                MaxTimeout = -1,
            };
            var client = new RestClient(options);
            var request = new RestRequest("/" + Realm + "/oauth2/token", Method.Post);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddParameter("grant_type", "client_credentials");
            request.AddParameter("client_id", "" + Client_Id + "@" + Realm + "");
            request.AddParameter("client_secret", Client_Secret);
            request.AddParameter("resource", Dynamics_Url);
            RestResponse response = client.Execute(request);



            if (response.Content != "" && response.StatusCode.ToString() == "OK")
            {
                var dyArray = JsonConvert.DeserializeObject<dynamic>(response.Content!)!;
                //var dyArray = JObject.Parse(response.Content!);
                AccessToken = dyArray.access_token;
                return AccessToken;

            }

            return null!;
        }

        #endregion         
    }

    
}