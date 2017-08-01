    public class AccessComm
    {
        //Connection string
        private string ConStr;

        public AccessComm(string dbfullfilepath)
        {
            ConStr = string.Format("Provider=Microsoft.ACE.Oledb.16.0; Data Source={0}", dbfullfilepath);
        }

        //Used to retrieve information from the database
        public DataTable SelectQuery(string querystring)
        {
            OleDbConnection connection = new OleDbConnection(ConStr);
            DataTable dt = new DataTable();
            
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(querystring, connection))
            {
                adapter.Fill(dt);
            }
            return dt;
        }



        //Used to delete, update, or insert into database
        public void NoReturnQuery(string query)
        {
            OleDbConnection connection = new OleDbConnection(ConStr);
            OleDbCommand oleDbCommand = new OleDbCommand(query, connection);
            try
            {
                connection.Open();
                oleDbCommand.ExecuteNonQuery();
            }
            catch (Exception excep)
            {
                Console.WriteLine(excep.Source);
            }
            finally
            {
                connection.Close();
            }
        }

    }
