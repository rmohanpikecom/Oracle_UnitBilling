using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ApplicationBlocks.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Dynamics_Oracle_UnitBilling
{
    public class clsDAL
    {
        public static string dynConnectionString = System.Configuration.ConfigurationManager.AppSettings["DynamicsConnection"]!.ToString();

        public static IConfigurationRoot Config()
        {
            var configuration = new ConfigurationBuilder()
               .AddJsonFile("DBQueries.json", optional: false, reloadOnChange: true)
               .AddEnvironmentVariables()
               .Build();
            return configuration;
        }

        #region 

        #region Dynamics_UnitBilling_BatchHeader_Create
        public static DataSet Dynamics_UnitBilling_BatchHeader_Create()
        {
            var configuration = Config();
            string BatchSize = System.Configuration.ConfigurationManager.AppSettings["BatchSize"]!.ToString();
            string dsn = dynConnectionString;
            string cmd = configuration.GetSection("qryUnitBillingBatchHeader_Create").Value!.Replace("[BatchSize]", BatchSize);
            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.Text, cmd);
            return ds;
        }
        #endregion

        #region Dynamics_UniBilling_GetList_Create
        public static DataSet Dynamics_UniBilling_List_Create(string BatchName)
        {
            var configuration = Config();
            string BatchSize = System.Configuration.ConfigurationManager.AppSettings["BatchSize"]!.ToString();
            string dsn = dynConnectionString;
            string cmd = configuration.GetSection("qryUnitBillingGetList_Create").Value!.Replace("@BatchName", "'" + BatchName + "'");
            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.Text, cmd);
            return ds;
        }
        #endregion



        #region Dynamics_UnitBilling_BatchHeader_Update
        public static DataSet Dynamics_UnitBilling_BatchHeader_Update()
        {
            var configuration = Config();
            string BatchSize = System.Configuration.ConfigurationManager.AppSettings["BatchSize"]!.ToString();
            string dsn = dynConnectionString;
            string cmd = configuration.GetSection("qryUnitBillingBatchHeader_Update").Value!.Replace("[BatchSize]", BatchSize);
            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.Text, cmd);
            return ds;
        }
        #endregion

        #region Dynamics_UniBilling_GetList_Update
        public static DataSet Dynamics_UniBilling_List_Update(string BatchName)
        {
            var configuration = Config();
            string BatchSize = System.Configuration.ConfigurationManager.AppSettings["BatchSize"]!.ToString();
            string dsn = dynConnectionString;
            string cmd = configuration.GetSection("qryUnitBillingGetList_Update").Value!.Replace("@BatchName", "'" + BatchName + "'");
            DataSet ds = SqlHelper.ExecuteDataset(dsn, CommandType.Text, cmd);
            return ds;
        }
        #endregion

        #endregion

    }
}
