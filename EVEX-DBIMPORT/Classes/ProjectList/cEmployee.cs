using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEX_DBIMPORT
{
    class cEmployee
    {
        public string name;
        public string surname;
        public string email;
        public string password;
        public string passkey;

        public cEmployee(string name, string surname, string email, string password, string passkey)
        {
            this.name = name;
            this.surname = surname;
            this.email = email;
            this.password = password;
            this.passkey = passkey;
        }


        public void addToTable()
        {
            MySqlCommand cmd = new MySqlCommand("INSERT INTO employee(Employee_id, Name, surname, email, password, passkey) VALUES (@id, @name, @surname, @email, @password, @passkey)", DatabaseConnection.getCon());
            cmd.Parameters.AddWithValue("@id", 0);
            cmd.Parameters.AddWithValue("@name", this.name);
            cmd.Parameters.AddWithValue("@surname", this.surname);
            cmd.Parameters.AddWithValue("@email", this.email);
            cmd.Parameters.AddWithValue("@password", this.password);
            cmd.Parameters.AddWithValue("@passkey", this.passkey);

            cmd.ExecuteNonQuery();
        }
    }
}
