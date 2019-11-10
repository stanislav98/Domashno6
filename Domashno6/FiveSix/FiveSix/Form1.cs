using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace FiveSix
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "The average temperature for the week was: ";
            button1.Text = "Temperature";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Button button2 = new Button();
                button2.Size = new System.Drawing.Size(200, 50);
                button2.Location = new System.Drawing.Point(350, 300);
                button2.Text = "Save-EXCEL";
                button2.Click += new EventHandler(this.button2_Click);
                Button button3 = new Button();
                button3.Size = new System.Drawing.Size(200, 50);
                button3.Location = new System.Drawing.Point(140, 300);
                button3.Text = "Save-WORD";
                button3.Click += new EventHandler(this.button3_Click);
                this.Controls.Add(button2);
                this.Controls.Add(button3);
                double sum = 0;
                string[] days = { "Sunday", "Monday", "TuesDay", "Wednesday", "Thursday", "Friday", "Saturday" };
                foreach (string day in days)
                {
                    double temperatura = (double.Parse(Interaction.InputBox("The temperature in " + day + " is: ")));
                    listBox1.Items.Add(temperatura + " degrees Celsius \nwas the temperature " + day);
                    sum = sum + temperatura;
                }
                sum = sum / 7;
                listBox1.Sorted = true;
                label1.Text = label1.Text + string.Format("{0:0.##}", (sum)) + " degrees centigrade."
                    + "\nThe minimum temperature = " + string.Format("{0:0.##}", listBox1.Items[0].ToString())
                    + "\nThe maximum temperature = " + string.Format("{0:0.##}", listBox1.Items[6].ToString());

            }
            catch
            {
                MessageBox.Show("Please input only numbers !");
                listBox1.Items.Clear();
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string fileName = @"D:\FiveSix.docx";
                var doc = DocX.Create(fileName);
                string title = "Docx Document Created Thanks To Nuget - > DocX ";
                Formatting titleFormat = new Formatting();
                titleFormat.FontFamily = new Xceed.Document.NET.Font("Batang");
                titleFormat.Size = 18;
                titleFormat.Position = 40;
                titleFormat.FontColor = System.Drawing.Color.Orange;
                titleFormat.UnderlineColor = System.Drawing.Color.Gray;
                titleFormat.Italic = true;
                Paragraph paragraphTitle = doc.InsertParagraph(title, false, titleFormat);
                paragraphTitle.Alignment = Alignment.center;
                foreach (string s in listBox1.Items)
                {
                    doc.InsertParagraph(s);
                }
                doc.InsertParagraph(label1.Text);
                doc.Save();
                MessageBox.Show("The file was created at D: and now will be started ! ");
                Process.Start("WINWORD.EXE", fileName);
            }
            catch
            {
                MessageBox.Show("Some error occurred");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add("FiveSixProject");
                    var excelWorksheet = excel.Workbook.Worksheets["FiveSixProject"];
                    string header = "Weather forecast for the week ! Made By Stanislav Stoychev thanks to NUGET - > EPPLUS";
                    string headerRange = "C1:F1";
                    excelWorksheet.Cells[headerRange].LoadFromText(header);
                    var start = excelWorksheet.Dimension.Start;
                    var end = excelWorksheet.Dimension.End;
                    int i = 1;
                    foreach (string str in listBox1.Items)
                    {
                        excelWorksheet.Column(i).Width = 40;
                        i++;
                        excelWorksheet.Cells[2, i].LoadFromText(str);
                    }
                    excelWorksheet.Cells[3, 1].LoadFromText(label1.Text);
                    FileInfo excelFile = new FileInfo(@"D:\FiveSix.xlsx");
                    excel.SaveAs(excelFile);
                    MessageBox.Show("Your File Was Created at D: and now will be opened ! ");
                    Process.Start("EXCEL.EXE", @"D:\FiveSix.xlsx");
                }
            }
            catch
            {
                MessageBox.Show("Some error occurred");
            }
           
        }
    }
}
