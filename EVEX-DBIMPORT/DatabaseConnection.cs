using MySql.Data;
using MySql;
using MySql.Data.MySqlClient;

namespace EVEX_DBIMPORT
{
    public class DatabaseConnection
    {
        private string ConnectionString = "Server=localhost;Database=evex_db;Uid=root;";

        private static MySqlConnection connection;

        public DatabaseConnection()
        {
            connection = new MySqlConnection(ConnectionString);
            connection.Open();
        }

        public bool isConnected()
        {
            if(connection.State == System.Data.ConnectionState.Open)
            {
                return true;
            }

            return false;
        }

        public static MySqlConnection getCon()
        {
            return connection;
        }
    }
}
