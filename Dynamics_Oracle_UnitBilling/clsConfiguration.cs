using System.Configuration;

namespace Dynamics_Oracle_UnitBilling
{
    public class clsConfiguration
    {
        public const string ConnectionString_Key = "ConnectionString";
        public const string Logconnectionstring_Key = "Logconnectionstring";
        public const string Encrypt_key = "Encrypt";
        public const string LogEncrypt_key = "LogEncrypt";

        private static volatile clsConfiguration _instance = null!;
        private string? strConnectString;
        private string? strLogConnectString;
        private string? strEncrypt;
        private string? strLogEncrypt;

        #region clsConfiguration
        private clsConfiguration()
        {
            strConnectString = ConfigurationManager.AppSettings[ConnectionString_Key];
            if (strConnectString == null)
                throw new ApplicationException(string.Format("No exist {0} in {1}.", ConnectionString_Key, ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath));

            strLogConnectString = ConfigurationManager.AppSettings[Logconnectionstring_Key];
            if (strConnectString == null)
                throw new ApplicationException(string.Format("No exist {0} in {1}.", Logconnectionstring_Key, ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath));


            strEncrypt = ConfigurationManager.AppSettings[Encrypt_key];
            if (strEncrypt == null)
                throw new ApplicationException(string.Format("No exist {0} in {1}.", Encrypt_key, ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath));

            strLogEncrypt = ConfigurationManager.AppSettings[LogEncrypt_key];
            if (strLogEncrypt == null)
                throw new ApplicationException(string.Format("No exist {0} in {1}.", LogEncrypt_key, ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath));

        }
        #endregion

        #region CurrentConfig
        public static clsConfiguration CurrentConfig
        {
            get
            {
                if (_instance == null)
                {
                    lock (typeof(clsConfiguration))
                    {
                        if (_instance == null)
                        {
                            _instance = new clsConfiguration();
                        }
                    }
                }
                return _instance;
            }
        }
        #endregion

        #region ConnectionString
        public string ConnectionString
        {
            get
            {
                return strConnectString!.Replace("**********", clsTools.Decrypt(Encrypt, true));
            }
        }
        #endregion

        #region LogConnectionString
        public string LogConnectionString
        {
            get
            {
                return strLogConnectString!.Replace("**********", clsTools.Decrypt(LogEncrypt, true));
            }
        }
        #endregion

        #region Encrypt
        public string Encrypt
        {
            get
            {
                return strEncrypt!;
            }
        }
        #endregion

        #region LogEncrypt
        public string LogEncrypt
        {
            get
            {
                return strLogEncrypt!;
            }
        }
        #endregion
    }
}
