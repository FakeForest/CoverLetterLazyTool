using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;

namespace CoverLetterLazyTool
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Word.Application wordApp;
        private Document document;

        public Form1()
        {
            InitializeComponent();
        }

        private void FindAndReplace(string searchText, string replacementText)
        {
            foreach (Range range in document.StoryRanges)
            {
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: searchText, ReplaceWith: replacementText, Replace: WdReplace.wdReplaceAll);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            wordApp = new Microsoft.Office.Interop.Word.Application();
            document = wordApp.Documents.Add();

            button1.Enabled = false;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;

            MessageBox.Show("Document created. You can start adding content.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (document != null)
            {
                if (textBox3.Text != "")
                {
                    string paragraphText = textBox3.Text;

                    // Insert the paragraph into the document
                    Paragraph para = document.Content.Paragraphs.Add();
                    para.Range.Text = paragraphText;

                    // Move to the next line for the next paragraph
                    document.Content.InsertParagraphAfter();

                    // Clear the textbox for the next input
                    //textBox3.Clear();

                    MessageBox.Show("Paragraph added successfully!");
                }
                else
                {
                    textBox3.Focus();
                    MessageBox.Show("Your paragraph added should not be empty!\n Please try again!");
                }
            }
            else
            {
                MessageBox.Show("Please create a document first.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (document != null)
            {
                if (textBox1.Text == "" || textBox2.Text == "")
                {
                    textBox1.Focus();
                    MessageBox.Show("Either the company name or the occupation is not supposed to be empty,\n Please try again!");
                }
                string searchText1 = "[CompanyName]";
                string replacementText1 = textBox1.Text;

                string searchText2 = "[Occupation]";
                string replacementText2 = textBox2.Text;

                FindAndReplace(searchText1, replacementText1);
                FindAndReplace(searchText2, replacementText2);

                MessageBox.Show("Word replaced successfully!");
            }
            else
            {
                MessageBox.Show("Please create a document first.");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (document != null)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    document.SaveAs2(filePath);
                    document.Close();
                    wordApp.Quit();
                    MessageBox.Show($"Document exported and saved to {filePath}");
                    button1.Enabled = true;
                }
            }
            else
            {
                MessageBox.Show("Please create a document first.");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }
    }
}
