using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleText
{
    public class ConnectionBase
    {
        static readonly string ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + AppDomain.CurrentDomain.BaseDirectory + @"DatabaseModel.mdf;Integrated Security=True";

        private IDbConnection Connection { get; set; }
        public ConnectionBase()
        {
            Connection = new SqlConnection(ConnectionString);
        }
        public IDbConnection ObterConexao()
        {
            if (Connection != null)
            {
                if (Connection.State == ConnectionState.Closed)
                {
                    if (string.IsNullOrEmpty(Connection.ConnectionString))
                        Connection = new SqlConnection(ConnectionString);
                    Connection.Open();
                }
                return Connection;
            }
            else
                throw new Exception();
        }
    }
}
