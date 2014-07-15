using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ionic.Zip;

namespace ParseGV
{
    public partial class GV_text : Form
    {
        public GV_text()
        {
            InitializeComponent();
        }
        bool wait = true;
        
        List<Tuple<string,string,string,string>> messages = new List<Tuple<string,string,string,string>>();
        
        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                ParseData();
            }
            catch (Exception err)
            {

                MessageBox.Show(err.Message);
            }
        }

        private void ParseData()
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Zip File | *.zip";

            open.ShowDialog();

            if (open.FileName == "")
            {
                return;
            }

            ZipFile file = ZipFile.Read(open.FileName);

            string path = Path.GetRandomFileName();
            string random_path = Path.Combine(Path.GetTempPath(), path);

            while (Directory.Exists(random_path))
            {
                path = Path.GetRandomFileName();
                random_path = Path.Combine(Path.GetTempPath(), path);
            }

            Directory.CreateDirectory(random_path);

            file.ExtractAll(random_path);

            var dirs = Directory.GetDirectories(random_path);

            //This is where files are. They are a bunch of html files with Text in them
            string path_of_data = dirs[0] + @"\Voice\Calls\";

            var files = Directory.GetFiles(path_of_data);
            List<string> files_filtered = (from x in files where x.Contains("Text") select x).ToList();
            var browser = new WebBrowser();
            //This for to wait for document to complete before handling
            browser.DocumentCompleted += browser_DocumentCompleted;
            foreach (var item in files_filtered)
            {
                var text = File.ReadAllText(item);

                browser.DocumentText = text;
                wait = true;
                var doc = browser.Document;

                while (wait)
                {
                    Application.DoEvents();
                }
                ParseData(browser);
            }
        }

        void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            wait = false;
        }

        private  void ParseData(WebBrowser sender)
        {
            foreach (HtmlElement item in ((WebBrowser)sender).Document.All)
            {
                if (item.GetAttribute("className") == "message")
                {
                    var date = item.All[0];
                    var number = item.All[2];
                    var message = item.All[4];

                    var tup = Tuple.Create(date.InnerText, number.GetAttribute("href"), number.InnerText, message.InnerText);

                    messages.Add(tup);
                
                }
                
            }
      
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                WriteCSV();
            }
            catch (Exception err)
            {

                MessageBox.Show(err.Message);
            }
        }

        private void WriteCSV()
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "CSV | *.csv";
            save.ShowDialog();

            if (save.FileName == "")
                return;

            using (StreamWriter outh = new StreamWriter(save.FileName))
            {
                var csv = new CsvHelper.CsvWriter(outh);
                csv.WriteField("Date");
                csv.WriteField("Telephone Number");
                csv.WriteField("Person");
                csv.WriteField("Message");
                csv.NextRecord();

                foreach (var item in messages)
                {
                    csv.WriteField(item.Item1);
                    csv.WriteField(item.Item2);
                    csv.WriteField(item.Item3);
                    csv.WriteField(item.Item4);
                    csv.NextRecord();
                }
            }
        }
    }
}
