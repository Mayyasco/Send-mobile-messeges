using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Net;

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void checkNodes(TreeNode node, int check)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.StateImageIndex = check;
                checkNodes(child, check);
            }

        }
        private void checkParent(TreeNode node)
        {
            if (node.Name == "Node0") return;
            else
            {
                int c = 0,c1=0;
                foreach (TreeNode child in node.Parent.Nodes)
                {
                    if (child.StateImageIndex == 0 ) c++;
                    if (child.StateImageIndex == 1) c1++;
                }
                if (c1 > 0) node.Parent.StateImageIndex = 1;
                else
                {
                    if (c == node.Parent.Nodes.Count) node.Parent.StateImageIndex = 0;
                    else if (c == 0) node.Parent.StateImageIndex = 2;
                    else node.Parent.StateImageIndex = 1;
                }
                checkParent(node.Parent);

            }
        }
        private void kill_excel()
        {
            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    theprocess.Kill();
                    return;
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //try
            //{
            treeView1.ExpandAll();
            
            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    MessageBox.Show("close all excel files");
                    this.Close();
                    return;
                }
            }
            dataGridView1.Rows.Clear();
            ApplicationClass app;
            app = new ApplicationClass();
            Workbook workBook1;
            workBook1 = app.Workbooks.Open(Directory.GetCurrentDirectory() + "\\names.xlsx",
        0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
            //---------------
            int x = int.Parse(((Range)workSheet1.Cells[2, 6]).Value2.ToString());
            if (x == 0) { MessageBox.Show("·« ÌÊÃœ ”Ã·« "); }
            else
            {
                int j = 1;
                for (int i = 1; i < x+1; i++)
                {
                    if (!((Range)workSheet1.Cells[i + 1, 3]).Value2.ToString().Contains("999999"))
                    {
                        dataGridView1.Rows.Add(1);
                        dataGridView1.Rows[j - 1].Cells[1].Value = ((Range)workSheet1.Cells[i + 1, 1]).Value2;
                        dataGridView1.Rows[j - 1].Cells[2].Value = ((Range)workSheet1.Cells[i + 1, 2]).Value2;
                        dataGridView1.Rows[j - 1].Cells[3].Value = ((Range)workSheet1.Cells[i + 1, 3]).Value2;
                        dataGridView1.Rows[j - 1].Cells[4].Value = ((Range)workSheet1.Cells[i + 1, 4]).Value2;
                        dataGridView1.Rows[j - 1].Cells[5].Value = ((Range)workSheet1.Cells[i + 1, 5]).Value2;
                        j++;
                    }
                }
            } 
            //---------------
            workBook1.Close(false, Directory.GetCurrentDirectory() + "\\names.xlsx", false);
            kill_excel();
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
                dataGridView1.Rows[j].Cells[0].Value = false;
        //}
        //catch (Exception ee)
        //{
        //    MessageBox.Show("·œÌﬂ „‘ﬂ·…");
        //    kill_excel();
        //    this.Close();
        //    return;

        //}
           
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.X > e.Node.Bounds.Left - e.Node.StateImageIndex || e.X < e.Node.Bounds.Left - (e.Node.StateImageIndex + 16)) return;
            else
            {
                if (e.Node.StateImageIndex == 0) e.Node.StateImageIndex = 2;
                else e.Node.StateImageIndex = 0;

                checkNodes(e.Node, e.Node.StateImageIndex);
                checkParent(e.Node);
            }
        }


        private void treeView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Space && treeView1.SelectedNode!=null)
            {
                if (treeView1.SelectedNode.StateImageIndex == 0) treeView1.SelectedNode.StateImageIndex = 2;
                else treeView1.SelectedNode.StateImageIndex = 0;

                checkNodes(treeView1.SelectedNode, treeView1.SelectedNode.StateImageIndex);
                checkParent(treeView1.SelectedNode);
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {   int x=textBox8.Text.Length;
            label1.Text=x.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {try{
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
                dataGridView1.Rows[j].Cells[0].Value = false;

            TreeNode eng=treeView1.Nodes["Node0"].Nodes["Node2"].Nodes["Node5"];
            TreeNode mng = treeView1.Nodes["Node0"].Nodes["Node2"].Nodes["Node3"];
            TreeNode teach = treeView1.Nodes["Node0"].Nodes["Node2"].Nodes["Node6"];

            TreeNode kt1 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node8"].Nodes["Node16"];
            TreeNode kt2 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node7"].Nodes["Node12"];
            TreeNode ka1 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Node2ff"].Nodes["Node10ff"];
            TreeNode ka2 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Nodeff"].Nodes["Node6ff"];
            
            TreeNode et1 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node8"].Nodes["Node15"];
            TreeNode et2 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node7"].Nodes["Node11"];
            TreeNode ea1 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Node2ff"].Nodes["Node9ff"];
            TreeNode ea2 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Nodeff"].Nodes["Node5ff"];
            
            TreeNode ct1 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node8"].Nodes["Node13"];
            TreeNode ct2 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node7"].Nodes["Node9"];
            TreeNode ca1 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Node2ff"].Nodes["Node7ff"];
            TreeNode ca2 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Nodeff"].Nodes["Node3ff"];
             
            TreeNode tt1 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node8"].Nodes["Node14"];
            TreeNode tt2 = treeView1.Nodes["Node0"].Nodes["Node1"].Nodes["Node7"].Nodes["Node10"];
            TreeNode ta1 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Node2ff"].Nodes["Node8ff"];
            TreeNode ta2 = treeView1.Nodes["Node0"].Nodes["Node55"].Nodes["Nodeff"].Nodes["Node4ff"];
            //MessageBox.Show(tt1.Text + " " + tt2.Text + " " + ta1.Text + " " + ta2.Text);
           for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {    
                 if (eng.StateImageIndex == 2&&dataGridView1.Rows[i].Cells[2].Value.ToString()=="„Â‰œ”")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (mng.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "«œ«—Ì")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                  else if (teach.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "„⁄·„")
                     dataGridView1.Rows[i].Cells[0].Value = true;

                 else if (kt1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
      && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "ﬂÂ—»«¡ «” ⁄„«·")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (kt2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
 && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "ﬂÂ—»«¡ «” ⁄„«·")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ka1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
 && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "ﬂÂ—»«¡ «” ⁄„«·")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ka2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
 && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "ﬂÂ—»«¡ «” ⁄„«·")
                     dataGridView1.Rows[i].Cells[0].Value = true;

                 else if (ct1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
 && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·Õ«”Ê»")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ct2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
   && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·Õ«”Ê»")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ca1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
 && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·Õ«”Ê»")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ca2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
 && dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·Õ«”Ê»")
                     dataGridView1.Rows[i].Cells[0].Value = true;

                 else if (tt1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·« ’«·« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (tt2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·« ’«·« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ta1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·« ’«·« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ta2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·« ’«·« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;

                 else if (et1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·ﬂ —Ê‰Ì« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (et2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "ÿ«·»"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·ﬂ —Ê‰Ì« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ea1.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·«Ê· À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·ﬂ —Ê‰Ì« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;
                 else if (ea2.StateImageIndex == 2 && dataGridView1.Rows[i].Cells[2].Value.ToString() == "Ê·Ì «„—"
&& dataGridView1.Rows[i].Cells[4].Value.ToString() == "«·À«‰Ì À«‰ÊÌ" && dataGridView1.Rows[i].Cells[5].Value.ToString() == "«·ﬂ —Ê‰Ì« ")
                     dataGridView1.Rows[i].Cells[0].Value = true;


                 
            }
        }
        catch (Exception ee)
        {
            MessageBox.Show("·œÌﬂ „‘ﬂ·…");

        }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try{
            string url = "http://smsservice.hadara.ps:4545/SMS.ashx/bulkservice/sessionvalue/getbalance/?apikey=773B1FB521B43305AFBE24FAFBB3B4DF&providerID=1";
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse res = (HttpWebResponse)req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            int s=responseString.LastIndexOf("<AvilabeBalance>");
            int end = responseString.LastIndexOf("</AvilabeBalance>");
            string s1 = responseString.Remove(end);
            string s2 = s1.Remove(0, s+16);
           textBox1.Text=s2 ;
       }
       catch (Exception ee)
       {
           MessageBox.Show("·œÌﬂ „‘ﬂ·…");
       }
        }

        private void button4_Click(object sender, EventArgs e)
        {//try{
            if (textBox8.Text.Trim().Length < 3) { MessageBox.Show("ÌÃ» «‰  ﬂÊ‰ «·—”«·… √ﬂÀ— „‰ 3 Õ—Ê›"); return; }
            int i=0,r=0;
            for ( i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if ((bool)dataGridView1.Rows[i].Cells[0].Value == true )
                { r++; }
            }
            if (r > 0)
            {
                progressBar1.Value = 0; progressBar1.Visible = true; progressBar1.Maximum = 100000;

                List<string> name = new List<string>();
                List<string> type = new List<string>();
                int t = 0;
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if ((bool)dataGridView1.Rows[i].Cells[0].Value == true )
                    {
                        string url = "http://smsservice.hadara.ps:4545/SMS.ashx/bulkservice/sessionvalue/sendmessage/?apikey=773B1FB521B43305AFBE24FAFBB3B4DF&to=0"+dataGridView1.Rows[i].Cells[3].Value.ToString()+"&msg="+textBox8.Text; 
                        HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                        HttpWebResponse res = (HttpWebResponse)req.GetResponse();
                        string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
                        //MessageBox.Show(responseString);
                        if (!responseString.Contains("1"))
                        {
                            name.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                            type.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());
                        }
                        t++;
                        progressBar1.Value += (int)100000 / r;
                    }
                    
                }
                 if (name.Count == 0) MessageBox.Show(" „ «·«—”«· »‰Ã«Õ ·ﬂ· „‰  „ «Œ Ì«—Â„ Ê⁄œœÂ„" + " = " + t.ToString());
                else
                {
                     string s = "⁄œœ „‰  „ «Œ Ì«—Â„" +" = "+ t.ToString()+ Environment.NewLine;
                     s = s + "⁄œœ „‰ ·„ Ì „ «·«—”«· ·Â„ »‰Ã«Õ" + " = " + name.Count.ToString() + " ÊÂ„" +" : "+ Environment.NewLine;
                    for (int y = 0; y < name.Count; y++)
                        s = s + type[y] + " : " + name[y] + Environment.NewLine;
                    Form f2 = new Form2();
                    f2.StartPosition = FormStartPosition.CenterScreen;
                    f2.Controls["textBox1"].Text = s;
                    f2.Controls["button1"].Select();
                    f2.ShowDialog();
                }
            }
            else MessageBox.Show("·„ Ì „ «·«—”«· ··ﬂ·");
            progressBar1.Value = 0;
            progressBar1.Visible = false;
        //}
        //catch (Exception ee)
        //{
        //    MessageBox.Show("·œÌﬂ „‘ﬂ·…");
        //    progressBar1.Value = 0;
        //    progressBar1.Visible = false;

        //}
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            kill_excel();
        }

        private void button2_Click(object sender, EventArgs e)
        {
             try{
            string url = "http://smsservice.hadara.ps:4545/SMS.ashx/bulkservice/sessionvalue/getbalance/?apikey=773B1FB521B43305AFBE24FAFBB3B4DF&providerID=2";
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse res = (HttpWebResponse)req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            int s=responseString.LastIndexOf("<AvilabeBalance>");
            int end = responseString.LastIndexOf("</AvilabeBalance>");
            string s1 = responseString.Remove(end);
            string s2 = s1.Remove(0, s+16);
           textBox2.Text=s2 ;
       }
       catch (Exception ee)
       {
           MessageBox.Show("·œÌﬂ „‘ﬂ·…");
       }
        }

        
        }


    }
