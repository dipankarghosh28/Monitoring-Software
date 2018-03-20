using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Threading;
namespace server
{
    public partial class Server : Form
    {
        delegate void SetTextCallback(string text);
        TcpListener listener;
        TcpClient client;
        NetworkStream ns;
        Thread t = null;

        public Server()
        {
            InitializeComponent();
            listener = new TcpListener(4949);
            listener.Start();
            client = listener.AcceptTcpClient();
            ns = client.GetStream();
            t = new Thread(DoWork);
            t.Start();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        string[] s = new string[20];
        string[] w = new string[20];
        int b= 0;
        int index,j = 0;

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            label1.Text = String.Format("Enter Data {0}", textBox1.Text);
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
        
            string r = textBox1.Text;
            string[] words = r.Split(' ');//.split Function is used to split the string into single characters and store in string array word.
            foreach (string word in words)
            {   
                //MessageBox.Show(word); 
                try
                {
                    s[index] = "" + int.Parse(word);
                    j = index;
                    byte[] byteTime = Encoding.ASCII.GetBytes(s[b]+" ");
                    //DialogResult dialogResult = MessageBox.Show("Should we continue to add Parameter Data ?", "Notification", MessageBoxButtons.YesNo);
                    //if (dialogResult == DialogResult.Yes)
                    //{
                        ns.Write(byteTime, 0, byteTime.Length);
                        b++;
                        index++;
                    //}
                    //else if (dialogResult == DialogResult.No)
                    //{
                      //  break;
                    //}
                }
                catch(Exception)
                {  
                    if(textBox1.Text=="")
                    MessageBox.Show("Empty,Kindly Enter the Data");
                    else
                    MessageBox.Show("Wrong Data Entered");

                }
        }
            textBox1.Text = String.Empty;
            richTextBox1.Text = String.Empty;
            textBox2.Text = String.Empty;
     }
  

        public void DoWork()
        {
            byte[] bytes = new byte[1024];
            while (true)
            {
                int bytesRead = ns.Read(bytes, 0, bytes.Length);
                this.SetText(Encoding.ASCII.GetString(bytes, 0, bytesRead));
                //MessageBox.Show(Encoding.ASCII.GetString(bytes, 0, bytesRead));
           
            }
        }
        int id;
        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.richTextBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
                
            }
            else
            {
                this.richTextBox1.Text = this.richTextBox1.Text + text;
                id =Convert.ToInt32(richTextBox1.Text);
                //richTextBox1.Text = String.Empty;
                if (id == 1)
                {
                    MessageBox.Show("The Sub System selected is CCC (Command Control Console) ");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of CCC are : 15 ";//The textbox on Server side shall display the no. of Parameters.
                    
                }
                else if (id == 2)
                {
                    MessageBox.Show("The Sub System selected is DSPU (Digital Signal Processing Unit) ");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of DSPU are : 4 ";
                }
                else if (id == 3)
                {
                    MessageBox.Show("The Sub System selected is EM ( Energy Meter) ");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of EM are : 2 ";
                }
                else if (id == 4)
                {
                    MessageBox.Show("The Sub System selected is GC (Gimbal Control) ");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of GC are : 10 ";
                }
                else if (id == 5)
                {
                    MessageBox.Show("The Sub System selected is GPS/DMC");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of GGPS/DMC are : 5 ";
                }
                else if (id == 6)
                {
                    MessageBox.Show("The Sub System selected is LS1 (TEA CO2) ");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of LS1 (TEA CO2) are : 14 ";
                }
                else if (id == 7)
                {
                    MessageBox.Show("The Sub System selected is LS2 (Nd :YAG)");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of LS2 (Nd : YAG) are : 13 ";
                }
                else if (id == 8)
                {
                    MessageBox.Show("The Sub System selected is WM (Wavelength Meter) ");// Message Box to display the Sub System Selected.
                    textBox2.Text = "No. of Parameters of WM are : 2 ";
                }
            }
            
        }

        private void button2_Click(object sender, EventArgs e)//Button which shall clear the textbox.
        {
            textBox1.Text = String.Empty;//
            MessageBox.Show("Enter the next Parameter in the Box above.");//Message Box shall prompt to enter the next Parameter.
        }

        private void button3_Click(object sender, EventArgs e)//Button is present to stop the debugging of the system.
        {
            System.Diagnostics.Debugger.Break();//Function to break debugging.
        }
   }
}


