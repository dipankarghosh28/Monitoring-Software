using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.IO; // File.Exists()
using System.Data.OleDb; // OleDbConnection, OleDbDataAdapter, OleDbCommandBuilder
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Client_4
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application xlexcel;
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        OleDbConnection conn;        //OleDbConnection
        OleDbDataAdapter adapter;   //Adapter has been created using OleDbDataAdapter. 
        OleDbDataAdapter adapter1; //Adapter 1 has been created.
        DataTable dtMain;
        private const int portNum = 4848; //Set the Port no. Here For eg. 4848 has been set.
        delegate void SetTextCallback(string text);
        string DBPath;
        TcpClient client;//Declaring it as a client side.
        NetworkStream ns;
        Thread t = null;
        private const string hostName = "localhost";//You can change the IP address of the system here. For eg. For the sams system "Localhost" Has been used.

        public Form1()
        {
            InitializeComponent();
            client = new TcpClient(hostName, portNum);//Declaring the hostname,Portnum.
            ns = client.GetStream();                 //Getstream function is utilised.
            t = new Thread(DoWork);
            t.Start();//A new thread is started via this function.

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//DatagridView Selection->Selects the whole row.
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//DatagridView Selection->Selects the whole row.
            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;//DatagridView Selection->Selects the whole row.
            DataGridViewCheckBoxColumn CheckboxColumn1 = new DataGridViewCheckBoxColumn();
            DataGridViewCheckBoxColumn CheckboxColumn2 = new DataGridViewCheckBoxColumn();
            DataGridViewCheckBoxColumn CheckboxColumn3 = new DataGridViewCheckBoxColumn();
            CheckBox chk = new CheckBox();
            dataGridView1.Columns.Add(CheckboxColumn1);
            dataGridView2.Columns.Add(CheckboxColumn2);
            dataGridView3.Columns.Add(CheckboxColumn3);

        }

        private void Form1_Load(object sender, EventArgs e)// This is the Load event which occurs once the form loads.
        {
            DBPath = Application.StartupPath + "\\New System Parameters.mdb";//This New System Parameters is our database file created.
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBPath);// connect to DB
            conn.Open();//Connection is opened with Server side.

            // create table "Table_1" if not exists
            // DO NOT USE SPACES IN TABLE AND COLUMNS NAMES TO PREVENT TROUBLES WITH SAVING, USE _
            // OLEDBCOMMANDBUILDER DON'T SUPPORT COLUMNS NAMES WITH SPACES
            try
            {
                using (OleDbCommand cmd = new OleDbCommand("CREATE TABLE [Table_1] ([id] COUNTER PRIMARY KEY, [text_column] MEMO, [int_column] INT);", conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex) { if (ex != null) ex = null; }


            using (DataTable dt = conn.GetSchema("Tables"))  // To get all tables from the Database we have specified.
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i].ItemArray[dt.Columns.IndexOf("TABLE_TYPE")].ToString() == "TABLE")
                    {
                        comboBox1.Items.Add(dt.Rows[i].ItemArray[dt.Columns.IndexOf("TABLE_NAME")].ToString());//Adds the table names in the comboBox1.
                        comboBox2.Items.Add(dt.Rows[i].ItemArray[dt.Columns.IndexOf("TABLE_NAME")].ToString());//Adds the table names in the comboBox2.
                        comboBox3.Items.Add(dt.Rows[i].ItemArray[dt.Columns.IndexOf("TABLE_NAME")].ToString());//Adds the table names in the comboBox3.
                    }
                }
            }
        }

        int m;//Integer m declared , this m variable is used for selecting which datagridview is connected with cell content click event of all Datagridviews.
        private void chkItems_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];

                if (chk.Value == chk.FalseValue || chk.Value == null)
                {
                    chk.Value = chk.TrueValue;
                    m = 0;
                    m = 1;
                    m = 2;
                }
                else
                {
                    chk.Value = chk.FalseValue;
                }


            }

        }
        int b;

        //Button Section.
        int id;
        private void button1_Click(object sender, EventArgs e)//This is the 1st OK Button in the form , used for selecting the table in DGV1.
        {
            m = 0;
            if (comboBox1.SelectedItem == null) return;
            adapter = new OleDbDataAdapter("SELECT * FROM [" + comboBox1.SelectedItem.ToString() + "]", conn);
            new OleDbCommandBuilder(adapter);
            dtMain = new DataTable();
            adapter.Fill(dtMain);
            dataGridView1.DataSource = dtMain;//DataGridView1 is connected to dtMain.

            if (comboBox1.SelectedItem.ToString() == "CCC (Command Control Console)")//Selects CCC Subsystem.
            {
                //MessageBox.Show("#1 ");//This messagebox can be made visible by removing comment '//'.
                id = 1;
            }
            if (comboBox1.SelectedItem.ToString() == "DSPU (Digital Signal Processing Unit)")//Selects DSPU Subsystem.
            {
                //MessageBox.Show("#2");
                id = 2;
            }
            if (comboBox1.SelectedItem.ToString() == "EM ( Energy Meter)")//Selects EM Subsystem.
            {
                MessageBox.Show("#3");
                id = 3;
            }
            if (comboBox1.SelectedItem.ToString() == "GC (Gimbal Control)")//Selects GC Subsystem.
            {
                //MessageBox.Show("#4");
                id = 4;
            }
            if (comboBox1.SelectedItem.ToString() == "GPS/DMC")//Selects GPS/DMC Subsystem.
            {
                //MessageBox.Show("#5");
                id = 5;
            }
            if (comboBox1.SelectedItem.ToString() == "LS1 (TEA : CO2)")//Selects LS1 (TEA CO2) Subsystem.
            {
                //MessageBox.Show("#6");
                id = 6;
            }
            if (comboBox1.SelectedItem.ToString() == "LS2 (Nd : YAG)")//Selects LS2 (Nd : YAG) Subsystem.
            {
                //MessageBox.Show("#7");
                id = 7;
            }
            if (comboBox1.SelectedItem.ToString() == "WM (Wavelength Meter)")//Selects WM Subsystem.
            {
                //MessageBox.Show("#8");
                id = 8;
            }
            b = 0;
        }

        private void button2_Click(object sender, EventArgs e)//This is the 2nd OK Button in the form , used for selecting the table in DGV2.
        {
            m = 1;
            if (comboBox2.SelectedItem == null) return;
            adapter1 = new OleDbDataAdapter("SELECT * FROM [" + comboBox2.SelectedItem.ToString() + "]", conn);
            new OleDbCommandBuilder(adapter1);
            dtMain = new DataTable();
            adapter1.Fill(dtMain);
            dataGridView2.DataSource = dtMain;

            if (comboBox2.SelectedItem.ToString() == "CCC (Command Control Console)")//Selects CCC Subsystem.
            {
                //MessageBox.Show("#1 ");
                id = 1;
            }
            if (comboBox2.SelectedItem.ToString() == "DSPU (Digital Signal Processing Unit)")//Selects DSPU Subsystem.
            {
                //MessageBox.Show("#2");
                id = 2;
            }
            if (comboBox2.SelectedItem.ToString() == "EM ( Energy Meter)")//Selects EM Subsystem.
            {
                MessageBox.Show("#3");
                id = 3;
            }
            if (comboBox2.SelectedItem.ToString() == "GC (Gimbal Control)")//Selects GC Subsystem.
            {
                //MessageBox.Show("#4");
                id = 4;
            }
            if (comboBox2.SelectedItem.ToString() == "GPS/DMC")//Selects GPS/DMC Subsystem.
            {
                //MessageBox.Show("#5");
                id = 5;
            }
            if (comboBox2.SelectedItem.ToString() == "LS1 (TEA CO2)")//Selects LS1 (TEA CO2) Subsystem.
            {
                //MessageBox.Show("#6");
                id = 6;
            }
            if (comboBox2.SelectedItem.ToString() == "LS2 (Nd :YAG)")//Selects LS2 (Nd : YAG) Subsystem.
            {
                //MessageBox.Show("#7");
                id = 7;
            }
            if (comboBox2.SelectedItem.ToString() == "WM (Wavelength Meter)")//Selects WM Subsystem.
            {
                //MessageBox.Show("#8");
                id = 8;
            }
            b = 0;
        }


        private void button3_Click(object sender, EventArgs e)//This is the 3rd OK Button in the form , used for selecting the table in DGV3.
        {
            m = 2;//m variable is set to 2.
            if (comboBox3.SelectedItem == null) return;
            adapter1 = new OleDbDataAdapter("SELECT * FROM [" + comboBox3.SelectedItem.ToString() + "]", conn);
            new OleDbCommandBuilder(adapter1);
            dtMain = new DataTable();
            adapter1.Fill(dtMain);
            dataGridView3.DataSource = dtMain;

            if (comboBox3.SelectedItem.ToString() == "CCC (Command Control Console)")//Selects CCC Subsystem.
            {
                //MessageBox.Show("#1 ");
                id = 1;
            }
            if (comboBox3.SelectedItem.ToString() == "DSPU (Digital Signal Processing Unit)")//Selects DSPU Subsystem.
            {
                MessageBox.Show("#2");
                id = 2;
            }
            if (comboBox3.SelectedItem.ToString() == "EM ( Energy Meter)")//Selects EM Subsystem.
            {
                MessageBox.Show("#3");
                id = 3;
            }
            if (comboBox3.SelectedItem.ToString() == "GC (Gimbal Control)")//Selects GC Subsystem.
            {
                //MessageBox.Show("#4");
                id = 4;
            }
            if (comboBox3.SelectedItem.ToString() == "GPS/DMC")//Selects GPS/DMC Subsystem.
            {
                //MessageBox.Show("#5");
                id = 5;
            }
            if (comboBox3.SelectedItem.ToString() == "LS1 (TEA CO2)")//Selects LS1 (TEA CO2) Subsystem.
            {
                //MessageBox.Show("#6");
                id = 6;
            }
            if (comboBox3.SelectedItem.ToString() == "LS2 (Nd : YAG)")//Selects LS2 (Nd : YAG) Subsystem.
            {
                //MessageBox.Show("#7");
                id = 7;
            }
            if (comboBox3.SelectedItem.ToString() == "WM (Wavelength Meter)")//Selects WM Subsystem.
            {
                //MessageBox.Show("#8");
                id = 8;
            }
            b = 0;
        }


        int i, j, r, p, q;//Declaring variables for the RowIndex and ColumnIndex for each DGV.

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)//Cell Content to be selected for DGV1
        {
            m = 0;
            if (id == 1 || id == 2 || id == 3 || id == 4 || id == 5 || id == 6 || id == 7 || id == 8)
            {

                i = 0;//Row Index is 0(In our case we need to show parameter data colum content).
                j = 3;//Column Index is 3.
                if (m == 0)//m integer 0 signifies 1st DataGridView.
                {
                    textBox1.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();//The selected cell content is displayed in TextBox 1.
                }
                q = dataGridView1.Rows.Count;//q variable is storing the count of no. of rows of 1st DataGridView.

            }

        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)//Cell content to be selected for DGV2
        {
            m = 1;
            //i = dataGridView2.CurrentCell.RowIndex;
            //j = dataGridView2.CurrentCell.ColumnIndex;
            if (id == 1 || id == 2 || id == 3 || id == 4 || id == 5 || id == 6 || id == 7 || id == 8)
            {
                m = 1;
                i = 0;
                j = 3;
                if (m == 1)
                {
                    textBox1.Text = dataGridView2.Rows[i].Cells[j].Value.ToString();//The selected cell content is displayed in TextBox 1.
                }
                r = dataGridView2.Rows.Count;//r variable is storing the count of no. of rows of 1st DataGridView.

            }
        }


        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)//Cell content to be selected for DGV3
        {
            m = 2;
            //i = dataGridView3.CurrentCell.RowIndex;
            //j = dataGridView3.CurrentCell.ColumnIndex;
            if (id == 1 || id == 2 || id == 3 || id == 4 || id == 5 || id == 6 || id == 7 || id == 8)
            {
                m = 2;
                i = 0;
                j = 3;
                if (m == 2)
                {
                    textBox1.Text = dataGridView3.Rows[i].Cells[j].Value.ToString();

                }
                p = dataGridView3.Rows.Count;
            }
        }


        public void DoWork()//This Functon DoWork() is a part of the Data transferral.
        {
            byte[] bytes = new byte[1024];
            while (true)
            {
                int bytesRead = ns.Read(bytes, 0, bytes.Length);
                this.SetText(Encoding.ASCII.GetString(bytes, 0, bytesRead));
                //MessageBox.Show(Encoding.ASCII.GetString(bytes, 0, bytesRead));
            }
        }

        private void SetText(string text)
        {
            if (m == 0)
            {
                if (this.textBox2.InvokeRequired)//InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.If these threads are different, it returns true.
                {
                    SetTextCallback d = new SetTextCallback(SetText);
                    this.Invoke(d, new object[] { text });
                }
                else
                {

                    this.textBox2.Text = this.textBox2.Text + text;
                }
            }
            if (m == 1)
            {
                if (this.textBox2.InvokeRequired)//InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.If these threads are different, it returns true.
                {
                    SetTextCallback d = new SetTextCallback(SetText);
                    this.Invoke(d, new object[] { text });
                }
                else
                {
                    this.textBox2.Text = this.textBox2.Text + text;
                }
            }
            if (m == 2)
            {
                if (this.textBox2.InvokeRequired)//InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.If these threads are different, it returns true.
                {
                    SetTextCallback d = new SetTextCallback(SetText);
                    this.Invoke(d, new object[] { text });
                }
                else
                {
                    this.textBox2.Text = this.textBox2.Text + text;
                }
            }
        }
        string[] x;
        private void button4_Click(object sender, EventArgs e) //This is the button to transfer data from the textbox to the Gridview.
        {
            string[] s;
            int b = 0;//Row
            int u = 0;
            int index = 0;


            do
            {
                if (id == 1 || id == 2 || id == 3 || id == 4 || id == 5 || id == 6 || id == 7 || id == 8)
                {

                    string r = textBox2.Text;//String r stores the string from Textbox 2.
                    string[] words = r.Split(' ');//.split function splits the string into characters and stores in string array words.
                    s = new string[20];
                    j = 2;//Column
                    //MessageBox.Show("Working");
                    try
                    {
                        foreach (string word in words)
                        {

                            s[index] = "" + int.Parse(word);
                            //MessageBox.Show(s[index]);
                            index++;
                            dtMain.Rows[b][j] = s[u];//dtMain signifies the selected DataGridView.
                            b++;
                            u++;
                        }


                        //dtMain.Rows[b][j] = s[u];
                        //u++;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("All rows are filled.");
                        textBox1.Text = String.Empty;
                        //d = 1;

                        //MessageBox.Show("Exit");
                        //textBox2.Text = String.Empty;
                    }
                    q--;
                    //MessageBox.Show("Exit");

                }
                //MessageBox.Show(q+"");
                //MessageBox.Show(p+"");
                //MessageBox.Show(r+"");
            } while (q == 0);
            textBox2.Text = String.Empty;
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)//This is for error check.
        {
            if (dtMain.Columns[e.ColumnIndex].DataType == typeof(Int64) ||
                dtMain.Columns[e.ColumnIndex].DataType == typeof(Int32) ||
                dtMain.Columns[e.ColumnIndex].DataType == typeof(Int16))
            {
                Rectangle rectColumn;
                rectColumn = dataGridView1.GetColumnDisplayRectangle(e.ColumnIndex, false);

                Rectangle rectRow;
                rectRow = dataGridView1.GetRowDisplayRectangle(e.RowIndex, false);
            }
        }


        private void button5_Click(object sender, EventArgs e)
        {
        String s = ""+id;
        byte[] byteTime = Encoding.ASCII.GetBytes(s);
        ns.Write(byteTime, 0, byteTime.Length);
        textBox1.Text = String.Empty;
        textBox2.Text = String.Empty;
        }

        //Excel Section started.
        private void button6_Click(object sender, EventArgs e)
        {
            
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            // Open a File
            //xlWorkBook = xlexcel.Workbooks.Open(@"C:\Windows\SysWOW64\config\systemprofile\Desktop\MyFile1.xlsx", 0, true, 5, "", "", true,Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {
                xlWorkBook = xlexcel.Workbooks.Open(@"C:\Users\DEEPANKAR\Desktop\VC2++\MyFile2.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
              
                if (id == 1)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[1,1] = "Joystick X Min";
                    xlWorkSheet.Cells[1,2] = "Joystick X Max";
                    xlWorkSheet.Cells[1,3] = "Joystick Y Min";
                    xlWorkSheet.Cells[1,4] = "Joystick Y Max";
                    xlWorkSheet.Cells[1,5] = "CO2";
                    xlWorkSheet.Cells[1,6] = "Co2 Laser Fire";
                    xlWorkSheet.Cells[1,7] = "Nd YAG";
                    xlWorkSheet.Cells[1,8] = "Nd:YAG Laser Fire";
                    xlWorkSheet.Cells[1,9] = "DOME";
                    xlWorkSheet.Cells[1,10] ="DOME";
                    xlWorkSheet.Cells[1,11] = "Energy Meter";
                    xlWorkSheet.Cells[1,12] = "Wavelength Meter";
                    xlWorkSheet.Cells[1,13] = "CCD Camera";
                    xlWorkSheet.Cells[1,14] = "IR Camera";
                    xlWorkSheet.Cells[1,15] = "DSPU";
                }
                else if (id == 2)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);//Worksheet 2 is updated.
                    xlWorkSheet.Cells[1,1] = "DSPU Unit";
                    xlWorkSheet.Cells[1,2] = "Intensity";
                    xlWorkSheet.Cells[1,3] = "Range";
                    xlWorkSheet.Cells[1,4] = "Detector Active";
                }
                else if (id == 3)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
                    xlWorkSheet.Cells[1,1] = "Energy Meter";
                    xlWorkSheet.Cells[1,2] = "Laser Energy";
                }
                else if (id == 4)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);
                    xlWorkSheet.Cells[1,1] = "Az Encoder";
                    xlWorkSheet.Cells[1,2] = "El Encoder";
                    xlWorkSheet.Cells[1,3] = "Az Gyro";
                    xlWorkSheet.Cells[1,4] = "El Gyror";
                    xlWorkSheet.Cells[1,5] = "DOME";
                    xlWorkSheet.Cells[1,6] = "Az Demand";
                    xlWorkSheet.Cells[1,7] = "El Demand";
                    xlWorkSheet.Cells[1,8] = "Az Servo";
                    xlWorkSheet.Cells[1,9] = "Ele Servo";
                    xlWorkSheet.Cells[1,10] = "Modes";
                }
                else if (id == 5)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);
                    xlWorkSheet.Cells[1,1] = "GPS";
                    xlWorkSheet.Cells[1,2] = "DMC";
                    xlWorkSheet.Cells[1,3] = "DMC(Direction)";
                    xlWorkSheet.Cells[1,4] = "DMC(Az Angle)";
                    xlWorkSheet.Cells[1,5] = "DMC(El Angle)";
                }
                else if (id == 6)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);
                    xlWorkSheet.Cells[1,1] = "Laser Fire PPS";
                    xlWorkSheet.Cells[1,2] = "Laser";
                    xlWorkSheet.Cells[1,3] = "Trigger Mode";
                    xlWorkSheet.Cells[1,4] = "Laser Wavelength";
                    xlWorkSheet.Cells[1,5] = "No. of Shots";
                    xlWorkSheet.Cells[1,6] = "Trigger Control";
                    xlWorkSheet.Cells[1,7] = "HV Enable";
                    xlWorkSheet.Cells[1,8] = "HV Value";
                    xlWorkSheet.Cells[1,9] = "Charge";
                    xlWorkSheet.Cells[1,10] = "Stop on ARC Enable";
                    xlWorkSheet.Cells[1,11] = "Laser Mode";
                    xlWorkSheet.Cells[1,12] = "Laser State";
                    xlWorkSheet.Cells[1,13] = "Halt Error Code";
                    xlWorkSheet.Cells[1,14] = "Info Error Code";
                   
                }
                else if (id == 7)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(7);
                    xlWorkSheet.Cells[1,1] = "Laser Fire";
                    xlWorkSheet.Cells[1,2] = "Flash Lamp Status";
                    xlWorkSheet.Cells[1,3] = "Q Switch Status";
                    xlWorkSheet.Cells[1,4] = "Q-Switch Burst Count";
                    xlWorkSheet.Cells[1,5] = "Interlock Conditions";
                    xlWorkSheet.Cells[1,6] = "Flash Lamp Voltage";
                    xlWorkSheet.Cells[1,7] = "Flash Lamp Frequency";
                    xlWorkSheet.Cells[1,8] = "Flash Lamp Power";
                    xlWorkSheet.Cells[1,9] = "Flash Lamp Energy";
                    xlWorkSheet.Cells[1,10] = "Q-Switch Mode";
                    xlWorkSheet.Cells[1,11] = "Operating Temp";
                    xlWorkSheet.Cells[1,12] = "Q Switch Single Shot";
                    xlWorkSheet.Cells[1,13] = "PPS";
                }
                else if (id == 8)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(8);
                    xlWorkSheet.Cells[1, 1] = "Wavelength Meter";
                    xlWorkSheet.Cells[1, 2] = "Current Wavelength Transmitted";
                }
                else
                {
                    MessageBox.Show("Some problem");
                }
            }
            catch(Exception)
            { MessageBox.Show("Check the Excel File,There might be some problem to it."); }
            
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string[] d;
            d = new string[20];
            int i1 = 0;
            int u = 1;
            string a = textBox2.Text;
            string[] words = a.Split(' ');
            int _lastRow = xlWorkSheet.Range["A" + xlWorkSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row + 1;
          
            //MessageBox.Show("Working");
             try
              {
                 if (id == 1)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[1,1] = "Joystick X Min";
                    xlWorkSheet.Cells[1,2] = "Joystick X Max";
                    xlWorkSheet.Cells[1,3] = "Joystick Y Min";
                    xlWorkSheet.Cells[1,4] = "Joystick Y Max";
                    xlWorkSheet.Cells[1,5] = "CO2";
                    xlWorkSheet.Cells[1,6] = "Co2 Laser Fire";
                    xlWorkSheet.Cells[1,7] = "Nd YAG";
                    xlWorkSheet.Cells[1,8] = "Nd:YAG Laser Fire";
                    xlWorkSheet.Cells[1,9] = "DOME";
                    xlWorkSheet.Cells[1,10] ="DOME";
                    xlWorkSheet.Cells[1,11] = "Energy Meter";
                    xlWorkSheet.Cells[1,12] = "Wavelength Meter";
                    xlWorkSheet.Cells[1,13] = "CCD Camera";
                    xlWorkSheet.Cells[1,14] = "IR Camera";
                    xlWorkSheet.Cells[1,15] = "DSPU";
                }
                else if (id == 2)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                    xlWorkSheet.Cells[1,1] = "DSPU Unit";
                    xlWorkSheet.Cells[1,2] = "Intensity";
                    xlWorkSheet.Cells[1,3] = "Range";
                    xlWorkSheet.Cells[1,4] = "Detector Active";
                }
                else if (id == 3)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
                    xlWorkSheet.Cells[1,1] = "Energy Meter";
                    xlWorkSheet.Cells[1,2] = "Laser Energy";
                }
                else if (id == 4)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);
                    xlWorkSheet.Cells[1,1] = "Az Encoder";
                    xlWorkSheet.Cells[1,2] = "El Encoder";
                    xlWorkSheet.Cells[1,3] = "Az Gyro";
                    xlWorkSheet.Cells[1,4] = "El Gyror";
                    xlWorkSheet.Cells[1,5] = "DOME";
                    xlWorkSheet.Cells[1,6] = "Az Demand";
                    xlWorkSheet.Cells[1,7] = "El Demand";
                    xlWorkSheet.Cells[1,8] = "Az Servo";
                    xlWorkSheet.Cells[1,9] = "Ele Servo";
                    xlWorkSheet.Cells[1,10] = "Modes";
                }
                else if (id == 5)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);
                    xlWorkSheet.Cells[1,1] = "GPS";
                    xlWorkSheet.Cells[1,2] = "DMC";
                    xlWorkSheet.Cells[1,3] = "DMC(Direction)";
                    xlWorkSheet.Cells[1,4] = "DMC(Az Angle)";
                    xlWorkSheet.Cells[1,5] = "DMC(El Angle)";
                }
                else if (id == 6)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);
                    xlWorkSheet.Cells[1,1] = "Laser Fire PPS";
                    xlWorkSheet.Cells[1,2] = "Laser";
                    xlWorkSheet.Cells[1,3] = "Trigger Mode";
                    xlWorkSheet.Cells[1,4] = "Laser Wavelength";
                    xlWorkSheet.Cells[1,5] = "No. of Shots";
                    xlWorkSheet.Cells[1,6] = "Trigger Control";
                    xlWorkSheet.Cells[1,7] = "HV Enable";
                    xlWorkSheet.Cells[1,8] = "HV Value";
                    xlWorkSheet.Cells[1,9] = "Charge";
                    xlWorkSheet.Cells[1,10] = "Stop on ARC Enable";
                    xlWorkSheet.Cells[1,11] = "Laser Mode";
                    xlWorkSheet.Cells[1,12] = "Laser State";
                    xlWorkSheet.Cells[1,13] = "Halt Error Code";
                    xlWorkSheet.Cells[1,14] = "Info Error Code";
                   
                }
                else if (id == 7)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(7);
                    xlWorkSheet.Cells[1,1] = "Laser Fire";
                    xlWorkSheet.Cells[1,2] = "Flash Lamp Status";
                    xlWorkSheet.Cells[1,3] = "Q Switch Status";
                    xlWorkSheet.Cells[1,4] = "Q-Switch Burst Count";
                    xlWorkSheet.Cells[1,5] = "Interlock Conditions";
                    xlWorkSheet.Cells[1,6] = "Flash Lamp Voltage";
                    xlWorkSheet.Cells[1,7] = "Flash Lamp Frequency";
                    xlWorkSheet.Cells[1,8] = "Flash Lamp Power";
                    xlWorkSheet.Cells[1,9] = "Flash Lamp Energy";
                    xlWorkSheet.Cells[1,10] = "Q-Switch Mode";
                    xlWorkSheet.Cells[1,11] = "Operating Temp";
                    xlWorkSheet.Cells[1,12] = "Q Switch Single Shot";
                    xlWorkSheet.Cells[1,13] = "PPS";
                }
                else if (id == 8)
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(8);
                    xlWorkSheet.Cells[1, 1] = "Wavelength Meter";
                    xlWorkSheet.Cells[1, 2] = "Current Wavelength Transmitted";
                }
                else
                {
                    MessageBox.Show("Some problem");
                }
                foreach (string word in words)
                    {
                        
                         d[i1] = "" + int.Parse(word);
                         xlWorkSheet.Cells[_lastRow, u] = d[i1];
                         u++;
                         MessageBox.Show(d[i1]);
                         i1++;
                      }
                  }
                    catch (Exception)
                    {
                        MessageBox.Show("All Done");
                    }
            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            xlWorkBook.Close(true, misValue, misValue);
            xlexcel.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlexcel);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
