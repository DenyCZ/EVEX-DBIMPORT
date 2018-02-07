using MySql.Data;
using MySql;
using MySql.Data.MySqlClient;

namespace EVEX_DBIMPORT
{
    class DatabaseConnection
    {
        private string ConnectionString = "Server=localhost;Database=evex_db;Uid=root;";

        private MySqlConnection connection;

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
    }
}
