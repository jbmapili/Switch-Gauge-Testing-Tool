using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using OpcRcw.Da;
using Microsoft.VisualBasic.FileIO;
using MainUserControl;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;

namespace PC_DCS_Tool_No3
{
    public partial class Form1 : Form
    {

        DxpSimpleAPI.DxpSimpleClass opc = new DxpSimpleAPI.DxpSimpleClass();
        List<string[]> list = new List<string[]>();
        List<string> tag_no = new List<string> { };
        List<string> ins1 = new List<string> { };
        List<string> ins2 = new List<string> { };
        List<string> amAdrs = new List<string> { };
        List<string> opAdrs = new List<string> { };
        List<string> btnTxt = new List<string> { };
        List<string> swAdrs = new List<string> { };
        List<string> doAdrs = new List<string> { };
        List<string> ansAdrs = new List<string> { };
        List<string> listSearch = new List<string> { };
        int a = -1;
        int b = -1;
        public Panel getPanel
        {
            get
            {
                return panel1;
            }
            set
            {
                panel1 = value;
            }

        }
        public Form1()
        { 
            InitializeComponent();
        }

        private void btnListRefresh_Click(object sender, EventArgs e)
        {
            cmbServerList.Items.Clear();
            string[] ServerNameArray;
            opc.EnumServerList(txtNode.Text, out ServerNameArray);

            for (int a = 0; a < ServerNameArray.Count<string>(); a++)
            {
                cmbServerList.Items.Add(ServerNameArray[a]);
            }
            cmbServerList.SelectedIndex = 0;
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            if (panel1.Controls.Count > 0) {
                if (opc.Connect(txtNode.Text, cmbServerList.Text))
                {
                    
                    btnListRefresh.Enabled = false;
                    btnDisconnect.Enabled = true;
                    btnConnect.Enabled = false;
                    for (int i = 0; i < panel1.Controls.Count; i++)
                    {
                        UserControl1 btn = panel1.Controls[i] as UserControl1;
                        string[] target = new string[1];

                        object[] val;
                        int[] nErrorArray;
                        short[] wQualityArray;
                        OpcRcw.Da.FILETIME[] fTimeArray;
                        if (btn.AmAddrTextbox != "")
                        {
                            target[0] = btn.AmAddrTextbox;
                            if (opc.Read(target, out val, out wQualityArray, out fTimeArray, out nErrorArray) == true)
                            {
                                string val0="";
                                try 
                                {
                                    val0 = val[0].ToString();
                                }
                                catch (Exception)
                                { 
                                    val0 = "error";
                                    Debug.WriteLine("Error reading " + btn.AmAddrTextbox + " value from the OPC-Server");
                                }
                                finally 
                                {
                                    Debug.WriteLine(btn.AmAddrTextbox + " " + val0);
                                    if (val0 == "True" || val0 == "-1")
                                    {
                                        btn.AmStatusText = "On";
                                    }
                                    else if (val0 == "False" || val0 == "0")
                                    {
                                        btn.AmStatusText = "Off";
                                    }
                                    else
                                    {
                                        btn.AmStatusText = "On/Off";
                                    }
                                }
                            }
                        }
                        if (btn.OpAddrTextbox != "")
                        {
                            target[0] = btn.OpAddrTextbox;
                            if (opc.Read(target, out val, out wQualityArray, out fTimeArray, out nErrorArray) == true)
                            {
                                string val1 = "";
                                try
                                {
                                    val1 = val[0].ToString();
                                }
                                catch (Exception)
                                {
                                    val1 = "error";
                                    Debug.WriteLine("Error reading " + btn.OpAddrTextbox + " value from the OPC-Server");
                                }
                                finally
                                {
                                    Debug.WriteLine(btn.OpAddrTextbox + " " + val1);
                                    if (val1 == "0")
                                    {
                                        btn.OPStatTextbox = "None";
                                    }
                                    else if (val1 == "1")
                                    {
                                        btn.OPStatTextbox = "Prohibition";
                                    }
                                    else if (val1 == "2")
                                    {
                                        btn.OPStatTextbox = "Maintenance";
                                    }
                                    else if (val1 == "3")
                                    {
                                        btn.OPStatTextbox = "Broke";
                                    }
                                    else
                                    {
                                        //return;
                                    }
                                }
                            }
                        }
                    }
                }
            }            
        }

        private void FileReadBtn_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                list.Clear();
                using (TextFieldParser parser = new TextFieldParser(openFileDialog1.FileName, Encoding.GetEncoding(932)))
                {
                    parser.Delimiters = new string[] { "," };
                    //bool st = false;
                    while (true)
                    {
                        string[] parts = parser.ReadFields();
                        if (parts == null)
                        {
                            break;
                        }
                        if (parts[95] == "1") 
                        { 
                            list.Add(parts);
                        }
                    }

                    panel1.Controls.Clear();
                    tag_no.Clear();
                    ins1.Clear();
                    ins2.Clear();
                    amAdrs.Clear();
                    opAdrs.Clear();
                    btnTxt.Clear();
                    swAdrs.Clear();
                    doAdrs.Clear();
                    ansAdrs.Clear();
                    if (list.Count > 0)
                    {
                        progressBar1.Visible = true;
                        progressBar1.Maximum = list.Count;
                        backgroundWorker1.RunWorkerAsync();        

                        txtSearch.Enabled = true;
                        button1.Enabled = true;
                    }
                    else
                    {
                        Label message = new Label();
                        message.Text = "There are no lists inside the file.";
                        message.Location = new Point(0, 222);
                        message.Width = 200;
                        panel1.Controls.Add(message);
                    }
                }
            }
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            if (opc.Disconnect())
            {
                btnConnect.Enabled = true;
                btnListRefresh.Enabled = true;
                btnDisconnect.Enabled = false;
            }
        }
        private string ReadValCopy(object oVal, int n, int nError)
        {
            if (nError != 0)
            {
                return "Read Error";
            }            
            return oVal.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtSearch.Enabled = false;
            button1.Enabled = false;
            progressBar1.Visible = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Regex reg = new Regex(txtSearch.Text.Replace(@"\",@"\\")
                                                .Replace("(", "\\(")
                                                .Replace("$", "\\$")
                                                .Replace("|", "\\|")
                                                .Replace("^", "\\^")
                                                .Replace("[", "\\[")
                                                .Replace(")", "\\)")
                                                .Replace("*","\\*"), RegexOptions.IgnoreCase);
            listSearch.Clear();
            for (int c = 0; c < list.Count; c++)
            {
                UserControl1 btn = panel1.Controls[c] as UserControl1;
                btn.TagTextboxColor = SystemColors.Control;
                btn.Intru1TextboxColor = SystemColors.Control;
                btn.Intru2TextboxColor = SystemColors.Control;
                btn.AmAddrTextboxColor = SystemColors.Control;
                btn.OpAddrTextboxColor = SystemColors.Control;
                for (int num = 0; num < 4; num++)
                {
                    btn.BtnUcs[num].AnsAddrTextboxColor = SystemColors.Control;
                    btn.BtnUcs[num].DoAddrTextboxColor = SystemColors.Control;
                    btn.BtnUcs[num].ButtonTextboxColor = SystemColors.Control;
                    btn.BtnUcs[num].SwAddrTextboxColor = SystemColors.Control;                    
                }
            }
            if (txtSearch.Text != "") 
            { 
                for (int c = 0; c < list.Count; c++)
                {
                    Match tag = reg.Match(tag_no[c]);
                    if (tag.Success)
                    {
                        UserControl1 btn = panel1.Controls[c] as UserControl1;
                        btn.TagTextboxColor = Color.Pink;
                        listSearch.Add(c.ToString());
                    }
                    Match insOne = reg.Match(ins1[c]);
                    if (insOne.Success)
                    {
                        UserControl1 btn = panel1.Controls[c] as UserControl1;
                        btn.Intru1TextboxColor = Color.Pink;
                        listSearch.Add(c.ToString());
                    }
                    Match insTwo = reg.Match(ins2[c]);
                    if (insTwo.Success)
                    {
                        UserControl1 btn = panel1.Controls[c] as UserControl1;
                        btn.Intru2TextboxColor = Color.Pink;
                        listSearch.Add(c.ToString());
                    }
                    Match am = reg.Match(amAdrs[c]);
                    if (am.Success)
                    {
                        UserControl1 btn = panel1.Controls[c] as UserControl1;
                        btn.AmAddrTextboxColor = Color.Pink;
                        listSearch.Add(c.ToString());
                    }
                    Match op = reg.Match(opAdrs[c]);
                    if (op.Success)
                    {
                        UserControl1 btn = panel1.Controls[c] as UserControl1;
                        btn.OpAddrTextboxColor = Color.Pink;
                        listSearch.Add(c.ToString());
                    }
                }
                for (int a = 0; a < btnTxt.Count; a++)
                {
                    Match bt = reg.Match(btnTxt[a]);
                    if (bt.Success)
                    {
                        int btI = (a % 4);
                        int btR = Convert.ToInt32(Math.Floor(a / 4.0));

                        UserControl1 btn = panel1.Controls[btR] as UserControl1;
                        btn.BtnUcs[btI].ButtonTextboxColor = Color.Pink;

                        listSearch.Add(btR.ToString());
                    }

                }
                for (int a = 0; a < swAdrs.Count; a++)
                {
                    Match sw = reg.Match(swAdrs[a]);
                    if (sw.Success)
                    {
                        int btI = (a % 4);
                        int btR = Convert.ToInt32(Math.Floor(a / 4.0));

                        UserControl1 btn = panel1.Controls[btR] as UserControl1;
                        btn.BtnUcs[btI].SwAddrTextboxColor = Color.Pink;

                        listSearch.Add(btR.ToString());
                    }
                }
                for (int a = 0; a < doAdrs.Count; a++)
                {
                    Match dO = reg.Match(doAdrs[a]);
                    if (dO.Success)
                    {
                        int btI = (a % 4);
                        int btR = Convert.ToInt32(Math.Floor(a / 4.0));

                        UserControl1 btn = panel1.Controls[btR] as UserControl1;
                        btn.BtnUcs[btI].DoAddrTextboxColor = Color.Pink;

                        listSearch.Add(btR.ToString());
                    }
                }
                for (int a = 0; a < ansAdrs.Count; a++)
                {
                    Match ans = reg.Match(ansAdrs[a]);
                    if (ans.Success)
                    {
                        int btI = (a % 4);
                        int btR = Convert.ToInt32(Math.Floor(a / 4.0));

                        UserControl1 btn = panel1.Controls[btR] as UserControl1;
                        btn.BtnUcs[btI].AnsAddrTextboxColor = Color.Pink;

                        listSearch.Add(btR.ToString());
                    }
                }
                if (listSearch.Count > 0)
                {
                    Debug.WriteLine(listSearch[0]);
                    panel1.VerticalScroll.Value = (Convert.ToInt32(listSearch[0]) * 222);
                    lblSearch.Text = "";
                }
                else
                {
                    lblSearch.Text = "No results found!";
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            search();
        }

        private void search()
        {
            if (txtSearch.Text != "")
            {
                if (a > -1)
                {
                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.TagTextboxColor = SystemColors.Control;
                    btn.Intru1TextboxColor = SystemColors.Control;
                    btn.Intru2TextboxColor = SystemColors.Control;
                    btn.AmAddrTextboxColor = SystemColors.Control;
                    btn.OpAddrTextboxColor = SystemColors.Control;
                    if (b > -1)
                    {
                        btn.BtnUcs[b].AnsAddrTextboxColor = SystemColors.Control;
                        btn.BtnUcs[b].DoAddrTextboxColor = SystemColors.Control;
                        btn.BtnUcs[b].SwAddrTextboxColor = SystemColors.Control;
                        btn.BtnUcs[b].ButtonTextboxColor = SystemColors.Control;
                    }                    
                }
                if (tag_no.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = tag_no.IndexOf(txtSearch.Text.ToUpper());
                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.TagTextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (ins1.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = ins1.IndexOf(txtSearch.Text.ToUpper());
                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.Intru1TextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (ins2.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = ins2.IndexOf(txtSearch.Text.ToUpper());
                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.Intru2TextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (amAdrs.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = amAdrs.IndexOf(txtSearch.Text.ToUpper());
                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.AmAddrTextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (opAdrs.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = opAdrs.IndexOf(txtSearch.Text.ToUpper());
                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.OpAddrTextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (btnTxt.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = btnTxt.IndexOf(txtSearch.Text.ToUpper());
                    b = ((a) % 4);
                    a = Convert.ToInt32(Math.Floor(a / 4.0));

                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.BtnUcs[b].ButtonTextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (swAdrs.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = swAdrs.IndexOf(txtSearch.Text.ToUpper());
                    b = ((a) % 4);
                    a = Convert.ToInt32(Math.Floor(a / 4.0));

                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.BtnUcs[b].SwAddrTextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (doAdrs.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = doAdrs.IndexOf(txtSearch.Text.ToUpper());
                    b = ((a) % 4);
                    a = Convert.ToInt32(Math.Floor(a / 4.0));

                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.BtnUcs[b].DoAddrTextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
                else if (ansAdrs.Contains(txtSearch.Text, StringComparer.CurrentCultureIgnoreCase))
                {
                    a = ansAdrs.IndexOf(txtSearch.Text.ToUpper());
                    b = ((a) % 4);
                    a = Convert.ToInt32(Math.Floor(a / 4.0));

                    UserControl1 btn = panel1.Controls[a] as UserControl1;
                    btn.BtnUcs[b].AnsAddrTextboxColor = Color.Pink;
                    panel1.VerticalScroll.Value = (a * 222);
                }
            }
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                search();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int a = 0; a < 10; a++)
            {
                (panel1.Controls[a] as UserControl1).Dispose();
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int j = 0; j < list.Count; j++)
            {
                Thread.Sleep(100);
                backgroundWorker1.ReportProgress(j);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int j = e.ProgressPercentage;

            UserControl1 uc = new UserControl1(opc);
            uc.Location = new Point(0, 222 * (j));

            panel1.Controls.Add(uc);
            uc.TagTextbox = list[j][0];
            uc.Intru1Textbox = list[j][1];
            uc.Intru2Textbox = list[j][2];
            uc.AmAddrTextbox = list[j][114];
            uc.OpAddrTextbox = list[j][119];
            progressBar1.Value = e.ProgressPercentage;
            tag_no.Add(list[j][0]);
            ins1.Add(list[j][1]);
            ins2.Add(list[j][2]);
            amAdrs.Add(list[j][114]);
            opAdrs.Add(list[j][119]);
            for (int num = 0; num < uc.BtnUcs.Length; num++)
            {
                uc.BtnUcs[num].ButtonTextbox = list[j][98 + num];
                uc.BtnUcs[num].SwAddrTextbox = list[j][102 + num];
                uc.BtnUcs[num].DoAddrTextbox = list[j][106 + num];
                uc.BtnUcs[num].AnsAddrTextbox = list[j][110 + num];
                btnTxt.Add(list[j][98 + num]);
                swAdrs.Add(list[j][102 + num]);
                doAdrs.Add(list[j][106 + num]);
                ansAdrs.Add(list[j][110 + num]);
            }
            if (j + 1 == list.Count)
            {
                progressBar1.Visible = false;
            }
        }
    }
}
