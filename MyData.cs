using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace LxWorkList
{
    static class MyData
    {
        public static OleDbConnection SolidOrganDB;
        public static OleDbConnection RSystemDBF;

        private static string liveDB = Properties.Settings.Default.SolidOrganDBLive;
        private static string testDB = Properties.Settings.Default.SolidOrganDBTest;
        private static string testRSDBF = Properties.Settings.Default.RenalSystemDBFlocal;
        private static string liveRSDBF = Properties.Settings.Default.RenalSytemDBF;

        /// <summary>
        /// Overloaded method to read data or the schema of a table from database
        /// </summary>
        /// <param name="queryString">Query string</param>
        /// <param name="dataBase">OleDBConnection object</param>
        /// <param name="isSchema">true = return schema, false = return data</param>
        /// <returns>DataTable</returns>
        public static DataTable ReadTableFromDB(string queryString, OleDbConnection dataBase, bool isSchema)
        {
            DataTable dataTable = new DataTable();

            if (isSchema)
            {
                OleDbCommand command = new OleDbCommand(queryString, dataBase);
                try
                {
                    command.Connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    dataTable = reader.GetSchemaTable();
                    reader.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally
                {
                    command.Connection.Close();
                }
                return dataTable;
            }

            else
            {
                OleDbDataAdapter dataAdaptor = new OleDbDataAdapter(queryString, dataBase);
                try
                {
                    dataBase.Open();
                    dataAdaptor.Fill(dataTable);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    dataBase.Close();
                    dataAdaptor.Dispose();
                    dataTable.Dispose();
                }
                return dataTable;
            }
        }

        /// <summary>
        /// Overloaded method to read data of a table from database
        /// </summary>
        /// <param name="queryString">Query string</param>
        /// <param name="dataBase">OleDBConnection object</param>
        /// <returns>DataTable</returns>
        public static DataTable ReadTableFromDB(string queryString, OleDbConnection dataBase)
        {
            DataTable dataTable = new DataTable();
            OleDbDataAdapter dataAdaptor = new OleDbDataAdapter(queryString, dataBase);
            try
            {
                dataBase.Open();
                dataAdaptor.Fill(dataTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dataBase.Close();
                dataAdaptor.Dispose();
                dataTable.Dispose();
            }
            return dataTable;
        }

        public static void ConnectDB() //toggle databases and test DB connection
        {
            if (File.Exists(@"c:\temp\test.txt"))   //toggle between TEST and LIVE dbs
            {
                SolidOrganDB = new OleDbConnection(testDB);
                RSystemDBF = new OleDbConnection(testRSDBF);
            }
            else
            {
                SolidOrganDB = new OleDbConnection(liveDB);
                RSystemDBF = new OleDbConnection(liveRSDBF);
            }

            try      // test connection
            {
                SolidOrganDB.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                SolidOrganDB.Close();
            }
        }

    }
}
