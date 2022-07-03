using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Aspose.Words;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Web;
using DotNetBrowser;
using System.Net;


namespace TestingAndReport
{
    public partial class Form1 : Form
    {
        int vulnum = 0;
        string first = "empty.docx";
        string second = "empty.docx";
        string third = "empty.docx";
        string fourth = "empty.docx";
        string five = "empty.docx";
        string six = "empty.docx";
        string seven = "empty.docx";
        string eight = "empty.docx";
        string nine = "empty.docx";
        string ten = "empty.docx";
        
        public Form1()
        {
            InitializeComponent();
            string footer = textBox6.Text;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            richTextBox1.Text = File.ReadAllText(@"Auto.txt");
            richTextBox1.SelectionBackColor = System.Drawing.Color.White;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox2.Text = File.ReadAllText(@"Manual.txt");
            richTextBox2.SelectionBackColor = System.Drawing.Color.White;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = File.AppendText(@"Auto.txt"))
            {
                sw.WriteLine(textBox1.Text);
                MessageBox.Show("Added !! Click View to see the updated");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = File.AppendText(@"Manual.txt"))
            {
                sw.WriteLine(textBox2.Text);
                MessageBox.Show("Added !! Click View to see the updated");
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            Process[] runingProcess = Process.GetProcesses();
            for (int i = 0; i < runingProcess.Length; i++)
            {
                // compare equivalent process by their name
                if (runingProcess[i].ProcessName == "winword")
                {
                    // kill  running process
                    runingProcess[i].Kill();
                    MessageBox.Show("Killed");
                }
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "domain#" + DateTime.Now.ToString("dd/MM/yyyy") + "#By:shiva";
            textBox2.Text = "domain#" + DateTime.Now.ToString("dd/MM/yyyy") + "#By: shiva";
            textBox7.Text = "domain#" + DateTime.Now.ToString("dd/MM/yyyy") + "#By: Anand";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string[] words = textBox3.Text.Split(',');
            foreach (string word in words)
            {
                int startindex = 0;
                while (startindex < richTextBox1.TextLength)
                {
                    int wordstartIndex = richTextBox1.Find(word, startindex, RichTextBoxFinds.None);
                    if (wordstartIndex != -1)
                    {
                        richTextBox1.SelectionStart = wordstartIndex;
                        richTextBox1.SelectionLength = word.Length + 20;
                        richTextBox1.SelectionBackColor = System.Drawing.Color.Yellow;
                        MessageBox.Show(richTextBox1.SelectedText);
                    }
                    else
                        break;
                    startindex += wordstartIndex + word.Length;
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string[] words = textBox4.Text.Split(',');
            foreach (string word in words)
            {
                int startindex = 0;
                while (startindex < richTextBox2.TextLength)
                {
                    int wordstartIndex = richTextBox2.Find(word, startindex, RichTextBoxFinds.None);
                    if (wordstartIndex != -1)
                    {
                        richTextBox2.SelectionStart = wordstartIndex;
                        richTextBox2.SelectionLength = word.Length + 20;
                        richTextBox2.SelectionBackColor = System.Drawing.Color.Yellow;
                        MessageBox.Show(richTextBox2.SelectedText);
                    }
                    else
                        break;
                    startindex += wordstartIndex + word.Length;
                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox2.Items.Add(listBox1.SelectedItem.ToString());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            listBox2.Items.Remove(listBox2.SelectedItem);
        }

        private void button7_Click(object sender, EventArgs e)
        {

            //  string one = listBox2.SelectedItem.ToString()+".docx";
            // string two = "2.docx";
            // For 2



            Aspose.Words.Document dstdoc = new Aspose.Words.Document(first);
            Aspose.Words.Document srcDoc = new Aspose.Words.Document(second);
            Aspose.Words.Document srcDoc1 = new Aspose.Words.Document(third);
            Aspose.Words.Document srcDoc2 = new Aspose.Words.Document(fourth);
            Aspose.Words.Document srcDoc3 = new Aspose.Words.Document(five);
            Aspose.Words.Document srcDoc4 = new Aspose.Words.Document(six);
            Aspose.Words.Document srcDoc5 = new Aspose.Words.Document(seven);
            Aspose.Words.Document srcDoc6 = new Aspose.Words.Document(eight);
            Aspose.Words.Document srcDoc7 = new Aspose.Words.Document(nine);

            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = false;
            if (first != "empty.docx")
            {
                dstdoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
                dstdoc.Save("output.docx");
            }
            if (second != "empty.docx")
            {
                srcDoc1.AppendDocument(dstdoc, ImportFormatMode.KeepSourceFormatting);
                srcDoc1.Save("output.docx");
            }
            if (third != "empty.docx")
            {
                srcDoc2.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
                srcDoc2.Save("output.docx");
            }
            if (fourth != "empty.docx")
            {
                srcDoc3.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);
                srcDoc3.Save("output.docx");
            }
            if (five != "empty.docx")
            {
                srcDoc4.AppendDocument(srcDoc3, ImportFormatMode.KeepSourceFormatting);
                srcDoc4.Save("output.docx");
            }
            if (six != "empty.docx")
            {
                srcDoc5.AppendDocument(srcDoc4, ImportFormatMode.KeepSourceFormatting);
                srcDoc5.Save("output.docx");
            }
            if (seven != "empty.docx")
            {
                srcDoc6.AppendDocument(srcDoc5, ImportFormatMode.KeepSourceFormatting);
                srcDoc6.Save("output.docx");
            }
            if (eight != "empty.docx")
            {
                srcDoc7.AppendDocument(srcDoc6, ImportFormatMode.KeepSourceFormatting);

                srcDoc7.Save("output.docx");

            }


        }

        private void button10_Click(object sender, EventArgs e)
        {


            foreach (object item in listBox2.Items)
            {

                vulnum++;

            }
            MessageBox.Show("Total Count is : " + vulnum);
            try
            {
                if (vulnum > 0)
                {
                    first = listBox2.Items[0].ToString() + ".docx";
                    MessageBox.Show("First is : " + first);
                }
                if (vulnum > 1)
                {
                    second = listBox2.Items[1].ToString() + ".docx";
                    MessageBox.Show("Second is : " + second);
                }
                if (vulnum > 2)
                {
                    third = listBox2.Items[2].ToString() + ".docx";
                    MessageBox.Show("Third is : " + third);
                }
                if (vulnum > 3)
                {

                    fourth = listBox2.Items[3].ToString() + ".docx";
                    MessageBox.Show("Fourth is : " + fourth);
                }
                if (vulnum > 4)
                {
                    five = listBox2.Items[4].ToString() + ".docx";
                    MessageBox.Show("Five is : " + five);
                }

                if (vulnum > 5)
                {

                    six = listBox2.Items[5].ToString() + ".docx";
                    MessageBox.Show("Six is : " + six);
                }
                if (vulnum > 6)
                {
                    seven = listBox2.Items[6].ToString() + ".docx";
                    MessageBox.Show("Seven is : " + seven);
                }
                if (vulnum > 7)
                {
                    eight = listBox2.Items[7].ToString() + ".docx";
                    MessageBox.Show("Eight is : " + eight);
                }
                if (vulnum > 8)
                {
                    nine = listBox2.Items[8].ToString() + ".docx";
                    MessageBox.Show("Nine is : " + nine);
                }
                if (vulnum > 9)
                {
                    ten = listBox2.Items[9].ToString() + ".docx";
                    MessageBox.Show("Ten is : " + ten);
                }



            }
            catch (Exception ex)
            {
                ex.ToString();
                MessageBox.Show(ex.ToString());
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            ReplaceText();
            ReplaceFooter();

        }

        private void button13_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Word.Application oWord;
            Microsoft.Office.Interop.Word.Paragraphs paragraphs = null;
            

            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = false; //to avoid displaying the Word Application
            object strDocName = @"C:\Users\Anand\source\repos\TestingAndReport\bin\Debug\netcoreapp3.1\output.docx";
            object objBool = false;
            object objNull = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Document oMyDoc = oWord.Documents.Open(ref strDocName, ref objBool, ref objBool, ref objBool, ref objNull,
                 ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull);

            object objReference = System.Reflection.Missing.Value;
           

            object objStart = 0;
            object objEnd = oMyDoc.Words.Count;

            if (oMyDoc.Sections.Count > 0)
            {
                try
                {
                    string keyword = oMyDoc.Paragraphs.First.ToString();
                    if (keyword.Contains("Evaluation Only. Created with Aspose.Words. Copyright 2003-2021 Aspose Pty Ltd"))
                        {
                        oMyDoc.Paragraphs.First.Range.Delete();
                        }
                   

                }
                catch (Exception)
                {
                }

            }
            oMyDoc.Save();

           

        }
       

        public void ReplaceText()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document("output.docx");
            // Find and replace text in the document
            doc.Range.Replace("url1", textBox5.Text, new Aspose.Words.Replacing.FindReplaceOptions(FindReplaceDirection.Forward));
            // Save the Word document


            doc.Save("output.docx");
        }
        public void ReplaceFooter()
        {
            // Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            //wordApp.Visible = false;
            // Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(@"C:\Users\Anand\source\repos\TestingAndReport\bin\Debug\netcoreapp3.1\output.docx");


            Microsoft.Office.Interop.Word.Application oWord;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = false; //to avoid displaying the Word Application
            object strDocName = @"C:\Users\Anand\source\repos\TestingAndReport\bin\Debug\netcoreapp3.1\output.docx";
            object objBool = false;
            object objNull = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Document oMyDoc = oWord.Documents.Open(ref strDocName, ref objBool, ref objBool, ref objBool, ref objNull,
                 ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull, ref objNull);

            object objReference = System.Reflection.Missing.Value;
            object objFoot = "this is footer";

            object objStart = 0;
            object objEnd = oMyDoc.Words.Count;

            if (oMyDoc.Sections.Count > 0)
            {
                try
                {
                    string footer = textBox6.Text;
                    oMyDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[2].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[2].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[3].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[3].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[4].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[4].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[5].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[5].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[6].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[6].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[7].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[7].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[8].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[8].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[9].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[9].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[10].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                    oMyDoc.Sections[10].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = footer;
                }
                catch (Exception)
                {
                }

            }
            oMyDoc.Save();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ReplaceFooter();
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            string footer = textBox6.Text;
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            
        }

        private void button16_Click(object sender, EventArgs e)
        {
            string[] words = textBox7.Text.Split(',');
            foreach (string word in words)
            {
                int startindex = 0;
                while (startindex < richTextBox6.TextLength)
                {
                    int wordstartIndex = richTextBox6.Find(word, startindex, RichTextBoxFinds.None);
                    if (wordstartIndex != -1)
                    {
                        richTextBox6.SelectionStart = wordstartIndex;
                        richTextBox6.SelectionLength = word.Length + 20;
                        richTextBox6.SelectionBackColor = System.Drawing.Color.Yellow;
                        MessageBox.Show(richTextBox6.SelectedText);
                    }
                    else
                        break;
                    startindex += wordstartIndex + word.Length;
                }
            }
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            richTextBox6.Text = File.ReadAllText(@"query.txt");
            richTextBox2.SelectionBackColor = System.Drawing.Color.White;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = File.AppendText(@"query.txt"))
            {
                sw.WriteLine(textBox8.Text);
                MessageBox.Show("Added !! Click View to see the updated");
            }
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void button18_Click(object sender, EventArgs e)
        {
            
            if (checkBox1.Checked==true)
                 {

                string my_web_link = "https://dynamic-applications.org/downloads/startup-product-manager/";
                Process.Start(new ProcessStartInfo { FileName = my_web_link, UseShellExecute = true });

               // string url = "https:///google.com";
              // Process.Start("chrome.exe", "https://google.com");

            }
        }
    }

    }



