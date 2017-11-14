using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace LxWorkList
{
    public partial class frmKidneyLongWaiters : Form
    {

        OleDbConnection SolidOrganDB;
        OleDbConnection RSystemDBF;
        string LiveDB = Properties.Settings.Default.SolidOrganDBLive;
        string TestDB = Properties.Settings.Default.SolidOrganDBTest;
        string TestRSDBF = Properties.Settings.Default.RenalSystemDBFlocal;
        string LiveRSDBF = Properties.Settings.Default.RenalSytemDBF;      

        public frmKidneyLongWaiters()
        {
            InitializeComponent();
            testDBconnection();
        }

        private void testDBconnection() //toggle databases and test DB connection
        {
            if (File.Exists(@"c:\temp\test.txt"))   //toggle between TEST and LIVE dbs
            {
                SolidOrganDB = new OleDbConnection(TestDB);
                RSystemDBF = new OleDbConnection(TestRSDBF);
            }
            else
            {
                SolidOrganDB = new OleDbConnection(LiveDB);
                RSystemDBF = new OleDbConnection(LiveRSDBF);
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

        private void frmKidneyLongWaiters_Load(object sender, EventArgs e)
        {
  
        }

        private void 
         
    }
}
