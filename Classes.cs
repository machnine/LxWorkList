using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace LxWorkList
{
    class WorkList
    {
        public int WorkListNumber { get; set; }
        public int Class { get; set; }
        public int KitType { get; set; }
        public string ID { get; set; }
        public string Treated { get; set; }
    }

    class WorkListItem : WorkList
    {
        public int Position { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public int RenalID { get; set; }
        public string SampleDate { get; set; }
        public string AdsorbOut { get; set; }
    }

    static class MyData
    {

       //static method to return a DataTable from a table in database
        public static DataTable ReadTableFromDB(string querystring, OleDbConnection database) 
        {
            OleDbDataAdapter dataAdaptor = new OleDbDataAdapter(querystring, database);
            DataTable dataTable = new DataTable();
            try
            {
                database.Open();
                dataAdaptor.Fill(dataTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                database.Close();
                dataAdaptor.Dispose();
                dataTable.Dispose();
            }
            return dataTable;
        }

    }
}
