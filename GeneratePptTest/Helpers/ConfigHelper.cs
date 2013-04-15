using Nini.Config;

namespace GeneratePptTest.Helpers
{
    /// <summary>
    /// Class used for configurations.
    /// </summary>
    public static class ConfigHelper
    {
        private static IConfigSource source;

        /// <summary>
        /// Config DataBase Section Constant.
        /// </summary>
        private const string ConfigDBSectionConstant = "Postgre DB";

        /// <summary>
        /// Connection String Format Constant.
        /// </summary>
        private const string ConnectionStringFormatConstant = "Server={0};Port={1};User Id={2};Database={3};Password={4};";

        /// <summary>
        /// MS Connection String Format Constant.
        /// </summary>
        private const string ConnectionStringFormatMssqlConstant = "server={0};database={1};Trusted_Connection=True;";

        /// <summary>
        /// Server Constant.
        /// </summary>
        private const string ServerConstant = "Server";

        /// <summary>
        /// Port Constant.
        /// </summary>
        private const string PortConstant = "Port";

        /// <summary>
        /// User ID Constant.
        /// </summary>
        private const string UserIDConstant = "User Id";

        /// <summary>
        /// Database Constant.
        /// </summary>
        private const string DatabaseConstant = "Database";

        /// <summary>
        /// Password Constant.
        /// </summary>
        private const string PasswordConstant = "Password";

        /// <summary>
        /// Constant which holds the path of the settings file.
        /// </summary>
        private const string SettingsPathConstant = @"Settings.ini";

        /// <summary>
        /// Config MS DataBase Section Constant.
        /// </summary>
        private const string ConfigMSDBSectionConstant = "MS DB";

        #region Public Methods

        /// <summary>
        /// Connection String Generator
        /// </summary>
        /// <returns>Postgre Sql DB connection string</returns>
        public static string GenerateConnectionString()
        {
            LoadConfig();

            // Set the config to the DB section of the INI file.
            IConfig config = source.Configs[ConfigDBSectionConstant];

            return string.Format(
                ConnectionStringFormatConstant,
                config.Get(ServerConstant),
                config.Get(PortConstant),
                config.Get(UserIDConstant),
                config.Get(DatabaseConstant),
                config.Get(PasswordConstant));
        }

        /// <summary>
        /// Connection String Generator
        /// </summary>
        /// <returns>Postgre Sql DB connection string</returns>
        public static string GenerateConnectionStringFromFile(string filePath)
        {
            LoadConfigFromFile(filePath);

            // Set the config to the DB section of the INI file.
            IConfig config = source.Configs[ConfigDBSectionConstant];

            return string.Format(
                ConnectionStringFormatConstant,
                config.Get(ServerConstant),
                config.Get(PortConstant),
                config.Get(UserIDConstant),
                config.Get(DatabaseConstant),
                config.Get(PasswordConstant));
        }

        /// <summary>
        /// Connection String Generator.
        /// </summary>
        /// <returns>MS Sql DB connection string.</returns>
        public static string GenerateConnectionStringMssql()
        {
            LoadConfig();

            // Set the config to the DB section of the INI file.
            IConfig config = source.Configs[ConfigMSDBSectionConstant];

            return string.Format(
                ConnectionStringFormatMssqlConstant,
                config.Get(ServerConstant),
                config.Get(DatabaseConstant));
        }

        /// <summary>
        /// Connection String Generator.
        /// </summary>
        /// <returns>MS Sql DB connection string.</returns>
        public static string GenerateConnectionStringMssqlFromFile(string filePath)
        {
            LoadConfigFromFile(filePath);

            // Set the config to the DB section of the INI file.
            IConfig config = source.Configs[ConfigMSDBSectionConstant];

            return string.Format(
                ConnectionStringFormatMssqlConstant,
                config.Get(ServerConstant),
                config.Get(DatabaseConstant));
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Load the configuration source file.
        /// </summary>
        private static void LoadConfig()
        {
            if (source == null)
            {
                source = new IniConfigSource(SettingsPathConstant);
            }
        }

        /// <summary>
        /// Load the configuration source file from specified path.
        /// </summary>
        private static void LoadConfigFromFile(string filePath)
        {
            if (source == null)
            {
                source = new IniConfigSource(filePath);
            }
        }

        #endregion Private Methods
    }
}
