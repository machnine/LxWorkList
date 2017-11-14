using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Collections.Generic;

namespace LxWorkList
{
    public partial class frmLxWorkListGen : Form
    {
        OleDbConnection SolidOrganDB;
        List<WorkList> ListNotTested = new List<WorkList>();
        List<WorkListItem> CurrentWorkList = new List<WorkListItem>();

        string LiveDB = Properties.Settings.Default.SolidOrganDBLive;
        string TestDB = Properties.Settings.Default.SolidOrganDBTest;
        string TestRSDBF = Properties.Settings.Default.RenalSystemDBFlocal;
        string LiveRSDBF = Properties.Settings.Default.RenalSytemDBF;
        string RSDBF = "";
        string OutputPath = "";
        string OutputfileName = "";
         
        public frmLxWorkListGen()
        {
            InitializeComponent();
            testDBconnection();     //toggle databases and test DB connection
        }

        private void testDBconnection() //toggle databases and test DB connection
        {
            if (File.Exists(@"c:\temp\test.txt"))   //toggle between TEST and LIVE dbs
            {
                SolidOrganDB = new OleDbConnection(TestDB);
                RSDBF = TestRSDBF;
                OutputPath = @"c:\temp\";
            }
            else
            {
                SolidOrganDB = new OleDbConnection(LiveDB);
                RSDBF = LiveRSDBF;
                OutputPath = Properties.Settings.Default.LiveWorkListPath;
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

        private void Form1_Load(object sender, EventArgs e)
        {
            getUntestedList();
        }

        private void getUntestedList()//get a list of untested runs
        {
            string query = "SELECT worklistNum, kittypeID, kitclassID FROM antibodyworklists WHERE ISNULL(tested)";
            DataTable dTable = new DataTable();
            dTable = MyData.ReadTableFromDB(query, SolidOrganDB);
            treeView1.Nodes.Add("Luminex Worklists");

            foreach (DataRow list in dTable.Rows)
            {
                ListNotTested.Add(new WorkList
                {
                    WorkListNumber = Int32.Parse(list["worklistnum"].ToString().Trim()),
                    KitType = Int32.Parse(list["kittypeID"].ToString().Trim()),
                    Class = Int32.Parse(list["kitclassID"].ToString().Trim())
                });
            }

            foreach (WorkList list in ListNotTested)
            {
                switch (list.Class)
                {
                    case 1:
                        list.ID = "LI-";
                        break;
                    case 2:
                        list.ID = "LII-";
                        break;
                    case 3:
                        list.ID = "LS-";
                        break;
                }

                list.ID += list.WorkListNumber.ToString();

                switch (list.KitType)
                {
                    case 1:
                        list.ID += " (LSM)";
                        list.Treated = "";
                        break;
                    case 2:
                        list.ID += " (PRA)";
                        list.Treated = "";
                        break;
                    case 3:
                        list.ID += " (SAB)";
                        list.Treated = "E";
                        break;
                }                            
                treeView1.Nodes[0].Nodes.Add(list.WorkListNumber.ToString(), list.ID);
             }

            treeView1.Nodes[0].ExpandAll();
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (this.treeView1.SelectedNode.Text.IndexOf("Luminex") < 0)
            {
                getCurrentWorklist(Int32.Parse(this.treeView1.SelectedNode.Name));
                string temppath = this.treeView1.SelectedNode.Text;
                OutputfileName = temppath.Substring(0, temppath.IndexOf("(") - 1) + ".txt";
            }

            displayCurrentWorklist();
        }

        private void getCurrentWorklist(int worklistnumber)   //get currently selected worklist
        {
            List<WorkListItem> listOfSamples = new List<WorkListItem>(); //list to contain a list of samples on a single run
            DataTable dtAbSera = new DataTable();
            DataTable dtPatientDBF = new DataTable();
            DataTable dtScrPref = new DataTable();

            string query = @"SELECT [position], ptnumber, serumdate FROM antibodysera WHERE worklistid ="
                             + worklistnumber + " ORDER BY [position]";
            dtAbSera = MyData.ReadTableFromDB(query, SolidOrganDB);    //fill datatable with all samples in a worklist

            dtPatientDBF = readPatientsDBFfromRS();   //fill datatable with all patient names and renal ID

            query = "SELECT Num, ReqAbsorbMxScr, ReqAbsorbAll FROM screeningprefs";
            dtScrPref = MyData.ReadTableFromDB(query, SolidOrganDB);   //fill datatable with all AO information
             
        
            //generate a list from AntibodySera for a given worklist
            foreach (DataRow sample in dtAbSera.Rows)
            {
                //ignore control wells and format sample dates
                string sampledate = sample["serumdate"].ToString();
                if (sampledate.Length > 10)
                    sampledate = sampledate.Substring(0, 10).Replace("/", ".");

                //generate the list of samples except AB serum
                listOfSamples.Add(new WorkListItem
                {
                    Position = Int32.Parse(sample["position"].ToString()),
                    WorkListNumber = worklistnumber, //pass from treeview node name
                    RenalID = Int32.Parse(sample["ptnumber"].ToString()),
                    SampleDate = sampledate
                });
            }

            
            foreach (WorkListItem sample in listOfSamples)
            {
                foreach (WorkList list in ListNotTested)            //give all available worklist properties to worklistitems
                {
                    if (sample.WorkListNumber == list.WorkListNumber)
                    {
                        sample.WorkListNumber = list.WorkListNumber;
                        sample.Class = list.Class;
                        sample.KitType = list.KitType;
                        sample.ID = list.ID;
                        sample.Treated = list.Treated;
                    }
                }

                //if there is no date (ie. controls)
                if(sample.SampleDate.Length < 10)       
                {
                    switch (sample.Position)
                    {
                        case 1:
                            sample.LastName = "BLANK"; 
                            sample.Treated = "";
                            break;
                        case 2:
                            sample.LastName = "POS CONTROL";
                            break;
                        case 3:
                            sample.LastName = "NEG CONTROL";
                            break;
                    }
                }
                
                //find the lastname that matches patient number from RList
                foreach (DataRow patient in dtPatientDBF.Rows)               
                {
                    if (sample.RenalID.ToString() == patient["Renal Number"].ToString().Trim())
                    {
                        sample.LastName = patient["Last Name"].ToString();
                        sample.FirstName = patient["First Name"].ToString();
                    }
                        
                }

                //update the worklist with serum treatment (ie. EDTA for SAB)
                              

                //check if a serum sample requires AO
                foreach (DataRow patient in dtScrPref.Rows)
                {
                    if (sample.RenalID.ToString() == patient["num"].ToString().Trim())
                    {
                        if (patient["ReqAbsorbAll"].ToString().ToUpper() == "Y")
                        {
                            sample.AdsorbOut = "+AO";
                        }
                        else if ((patient["ReqAbsorbMxScr"].ToString().ToUpper() == "Y") && (sample.Class == 3))
                        {
                            sample.AdsorbOut = "+AO";
                        }
                    }
                    
                }
            }
            CurrentWorkList.Clear();
            CurrentWorkList = listOfSamples;
            dtAbSera.Dispose();
            dtPatientDBF.Dispose();
            dtScrPref.Dispose();                      
        }      

        private string outputPatientList()      //output patient list as .txt file for Luminex IS100 software import
        {
            string tempsampleid = "";       // temp string sample id
            string temptxt = "";       // temp string for Luminex machine list
            string tempcsv = "";       // temp string for Fusion patient list

            if (OutputfileName.IndexOf(".txt") > 0) //.txt exists = treeview nodes of runs have been clicked
            {
                StreamWriter txtWriter = new StreamWriter(OutputPath + OutputfileName);
                StreamWriter csvWriter = new StreamWriter(OutputPath + OutputfileName.Replace(".txt", ".csv"));

                txtWriter.WriteLine("LX100IS Patient List");               //default line 1
                txtWriter.WriteLine("[Accession#, Dilution Factor]");      //default line 2   ...didn't wrote in one Writeline so \r\n vs \n can be avoided?
               
                foreach (WorkListItem sample in CurrentWorkList)
                {
                    tempsampleid = sample.LastName + HSCTSuffix(sample) +
                                   sample.RenalID.ToString() + " " +
                                   sample.SampleDate + " " +
                                   sample.Treated +
                                   sample.AdsorbOut;
                    temptxt = tempsampleid.Replace(" -1", "");     //remove "-1" for control samples
                    temptxt = temptxt.Replace("  ", " ");     //remove excess white space

                    txtWriter.WriteLine(temptxt);


                    if (sample.RenalID >= 1)
                    {
                        tempcsv = "," + tempsampleid + "," + sample.RenalID + "," + sample.SampleDate;
                        csvWriter.WriteLine(tempcsv);
                    }

                    
                }
                txtWriter.WriteLine("LX AB SERUM");
                txtWriter.Close();
                csvWriter.Close();
                return "A patient list is exported to: " + (OutputPath + OutputfileName).ToUpper();
            }
            else
                return "Error making patient list!";
        }


         private string HSCTSuffix(WorkListItem sample)  //check names for "HSC(T)" 
        {
            string lastname, firstname;
           
            try
            {
                lastname = sample.LastName.ToUpper();                
            }
            catch
            {
                lastname = "";
            }

            try
            {
                firstname = sample.FirstName.ToUpper();
            }
            catch
            {
                firstname = "";
            }
            if ((lastname.IndexOf("HSC") != -1) || 
                (firstname.IndexOf("HSC")!= -1))
            {
                return " HSCT ";
            }
            else
            {
                return " ";
            }
        }


        private void btMakeList_Click(object sender, EventArgs e)
        {
            MessageBox.Show(outputPatientList());       //show where the patient list is exported
        }

        private void displayCurrentWorklist()  //get data, format and display in a datagrid 
        {
            DataTable cWorklist = new DataTable();      //a new data table to store only the items to be displayed columns[]
            int i = 0;
            string[] columnNames = new string[] { "Position", "Last Name", "First Name", "Renal ID", "Sample Date", "A/O"};

            for (i = 0; i < columnNames.Length; i++)        //make these columns
            {
                cWorklist.Columns.Add(columnNames[i]);
            }

            for (i = 0; i < CurrentWorkList.Count; i++)     //add items from Currentworklist to the display list
            {
                cWorklist.Rows.Add(
                    CurrentWorkList[i].Position,
                    CurrentWorkList[i].LastName,
                    CurrentWorkList[i].FirstName,
                    CurrentWorkList[i].RenalID.ToString().Replace("-1", ""),
                    CurrentWorkList[i].SampleDate,
                    CurrentWorkList[i].AdsorbOut
                    );
            }

            cWorklist.Rows.Add(i + 1, "LX AB SERUM");               //add AB serum to the list

            //code below set up the width of the datagridview
            this.dataGridView1.DataSource = cWorklist;  
            this.dataGridView1.Columns["Position"].Width = 60;
            this.dataGridView1.Columns["Last Name"].Width = 120;
            this.dataGridView1.Columns["First Name"].Width = 120;
            this.dataGridView1.Columns["Renal ID"].Width = 80;
            this.dataGridView1.Columns["Sample Date"].Width = 100;
            this.dataGridView1.Columns["A/O"].Width = 50;
        }

        private DataTable readPatientsDBFfromRS()      
        //read patients table from Renal system (can't read Rlist.dbf same error as importing into Access)
        {
            DataTable dtable = new DataTable();
            string[] colnames = new string[] { "Renal Number", "Last Name", "First Name" };
            for (int i = 0; i < colnames.Length; i++)
                dtable.Columns.Add(colnames[i]);
            OleDbConnection patientsDBF = new OleDbConnection(RSDBF);
            string query = "SELECT [number], last, first FROM patients";
            OleDbCommand command = new OleDbCommand(query, patientsDBF);
            patientsDBF.Open();
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                Int32 _number = Int32.Parse(reader["number"].ToString());
                String _last = reader["last"].ToString().Trim();
                String _first = reader["first"].ToString().Trim();

                dtable.Rows.Add(
                    _number, _last, _first
                    );


            }

            return dtable;
        }
    }
}

