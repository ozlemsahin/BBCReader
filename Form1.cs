using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using HTML = HtmlAgilityPack;
using System.Text.RegularExpressions;

namespace WindowsFormsApp6_ExtractFromWebInsertWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        Word.Application app;
        Word.Document worddoc;
        private void button1_Click_1(object sender, EventArgs e)
        {
            //open word document
            string path = @"C:\Users\saz5ib\Desktop\" + textBox1.Text;
            app = new Word.Application();
            worddoc = app.Documents.Open(path);
            app.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // get text from url
            string url = textBox2.Text;
            HTML.HtmlWeb web = new HTML.HtmlWeb();
            HTML.HtmlDocument webdoc = web.Load(url);
            string xpath = "";
            string newS = "";

            //regex
            Regex rx = new Regex(".*(&#x27;).*");
            Regex rx2 = new Regex(".*(&quot;).*");

            //get every path with text
            for (int var = 1; var < 50; var++)
            {
                xpath = "//*[@id='main-wrapper']/div/div/div[1]/main/div[" + var.ToString() + "]/p";
                var node = webdoc.DocumentNode.SelectSingleNode(xpath);


                //if this path is what we wanted
                if (node != null)
                {
                    //split words
                    string s = node.InnerText.ToString();
                    string[] arr = s.Split(' ');
                    string wordNew;

                    //change words with true version
                    foreach (string word in arr)
                    {
                        var match = rx.Match(word);
                        var match2 = rx2.Match(word);

                        if (match.Length != 0)
                        {
                            string[] miniarr = word.Split("&#x27;");
                            wordNew = miniarr[0] + "\'" + miniarr[1];
                            newS = newS + ' ' + wordNew;
                        }
                        else if (match2.Length != 0)
                        {
                            string[] miniarr = word.Split("&quot;");
                            wordNew = miniarr[0] + miniarr[1];
                            newS = newS + ' ' + wordNew;
                        }
                        else
                        {
                            newS = newS + ' ' + word;
                        }

                    }

                    //write to word document
                    worddoc.Content.InsertAfter(newS);

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            worddoc.Save();
            worddoc.Close();
            app.Quit();
            app.Visible = false;
        }
    }
}
