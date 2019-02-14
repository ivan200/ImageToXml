using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.IO;

namespace bitmapToTxt
{
    public partial class Form1 : Form
    {
        FileInfo openBmpFile;
        FileInfo saveTxtFile;

        Dictionary<string, int> colors = new Dictionary<string, int>();

        public Form1()
        {
            InitializeComponent();

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string fileName in files)
            {
                var file = new FileInfo(fileName);
                if (file.Exists) {
                    textBox1.Text = file.FullName;
                    return;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK && openFileDialog1.CheckFileExists) {
                textBox1.Text = openFileDialog1.FileName;
                openBmpFile = new FileInfo(openFileDialog1.FileName);
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog(this) == DialogResult.OK && saveFileDialog1.FileName.Length > 0) {
                textBox2.Text = saveFileDialog1.FileName;
                saveTxtFile = new FileInfo(saveFileDialog1.FileName);
            }
        }

        private string getColor(Color color) {
            if (color.A.ToString("X2") == "00") {
                return "FFFFFF";
            }

            return color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        }


        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0) {
                openBmpFile = new FileInfo(textBox1.Text);
            }

            if (!openBmpFile.Exists)
            {
                MessageBox.Show("File does not exist\n\n" + textBox1.Text, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (textBox2.Text.Length > 0)
            {
                saveTxtFile = new FileInfo(textBox2.Text);
            }

            if(textBox2.Text.Length == 0 || !saveTxtFile.Directory.Exists)
            {
                MessageBox.Show("Folder does not exist\n\n" + textBox2.Text, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            FileStream fs = null;
            bool error = false;
            try
            {
                fs = new FileStream(openBmpFile.FullName, FileMode.Open, FileAccess.Read);
                using (var bmp = (Bitmap)Image.FromStream(fs))
                {
                    Color color;
                    colors.Clear();
                    for (int y = 0; y < bmp.Height; y++)
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            color = bmp.GetPixel(x, y);
                            String strColor = getColor(color);

                            if (!colors.ContainsKey(strColor))
                            {
                                colors.Add(strColor, colors.Count);
                            }
                        }
                    }

                    StreamWriter steamWriter = new StreamWriter(saveTxtFile.FullName);

                    //Header
                    steamWriter.Write("<?xml version=\"1.0\"?>"
                        + "\r\n<?mso-application progid=\"Excel.Sheet\"?>"
                        + "\r\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\""
                        + "\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\""
                        + "\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\""
                        + "\r\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\""
                        + "\r\n xmlns:html=\"http://www.w3.org/TR/REC-html40\">");

                    //Styles
                    steamWriter.Write("\r\n <Styles>");
                    foreach (KeyValuePair<string, int> entry in colors)
                    {
                        steamWriter.Write(
                            "\r\n  <Style ss:ID=\"s" + entry.Value.ToString() + "\">"
                            + "\r\n   <Interior ss:Color=\"#" + entry.Key + "\" ss:Pattern=\"Solid\"/>"
                            + "\r\n  </Style>");
                    }
                    steamWriter.Write("\r\n </Styles>");


                    //Worksheet
                    steamWriter.Write("\r\n <Worksheet ss:Name=\"Книга1\">");

                    //Table
                    steamWriter.Write("\r\n  <Table ss:ExpandedColumnCount=\"" + bmp.Width +
                        "\" ss:ExpandedRowCount=\"" + bmp.Height
                        + "\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultColumnWidth=\"7.5\" ss:DefaultRowHeight=\"7.5\">");
                    for (int y = 0; y < bmp.Height; y++)
                    {
                        steamWriter.Write("\r\n   <Row ss:AutoFitHeight=\"0\">");
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            color = bmp.GetPixel(x, y);
                            String strColor = getColor(color);

                            steamWriter.Write("\r\n    <Cell ss:StyleID=\"s" + colors[strColor] + "\"/>");
                        }
                        steamWriter.Write("\r\n   </Row>");
                    }
                    steamWriter.Write("\r\n  </Table>");
                    steamWriter.Write("\r\n </Worksheet>");
                    steamWriter.Write("\r\n</Workbook>");
                    steamWriter.Close();
                }
            }
            catch (Exception ex) {
                error = true;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
            }
            finally
            {
                fs.Close();
            }
            if (!error) {
                MessageBox.Show("Converted!!", "OK", MessageBoxButtons.OK);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show(
                "Program for converting image to xml by pixels for opening in excel"
                + "\n\nMade by wilerat"
                + "\nusing MS Visual C# 2015"
                + "\nmail: wilerat@gmail.com"
                + "\nskype: Ivan.1010",
                "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var file = new FileInfo(textBox1.Text);
            if (file.Exists) {
                String newSaveFileName = file.Directory + "\\" + file.Name.Substring(0, file.Name.Length - file.Extension.Length) + ".xml";
                textBox2.Text = newSaveFileName;
                saveTxtFile = new FileInfo(newSaveFileName);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
