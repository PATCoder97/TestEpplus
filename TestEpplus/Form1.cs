
using OfficeOpenXml;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace TestEpplus
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string sourceFile = @"C:\Users\ANHTUAN\Desktop\VMI_LabCert\Test\CalLThuocCap.xlsx";


        private void button1_Click(object sender, EventArgs e)
        {
            string destFolder = @"C:\Users\ANHTUAN\Desktop\VMI_LabCert\Test\ABC";
            if (!Directory.Exists(destFolder)) Directory.CreateDirectory(destFolder);
            string destFile = Path.Combine(destFolder, $"{DateTime.Now:yyMMddhhmmss}.xlsx");
            string destFilePDF = Path.Combine(destFolder, $"{DateTime.Now:yyMMddhhmmss}.pdf");

            File.Copy(sourceFile, destFile, true);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            FileInfo newFile = new FileInfo(destFile);

            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets["BienBan"];
            //var ws = pck.Workbook.Worksheets.Add("BienBan");
            // ws.View.ShowGridLines = false;


            ws.InsertRow(37, 5);

            var sourceRange = ws.Cells["E36:X36"];
            for (int i = 0; i < 5; i++)
            {
                var destinationRange = ws.Cells[$"E{37 + i}:X{37 + i}"];
                sourceRange.Copy(destinationRange);
            }




            //ws.Column(4).OutlineLevel = 1;
            //ws.Column(4).Collapsed = true;
            //ws.Column(5).OutlineLevel = 1;
            //ws.Column(5).Collapsed = true;
            //ws.OutLineSummaryRight = true;

            ////Headers
            //ws.Cells["B1"].Value = "Name";
            //ws.Cells["C1"].Value = "Size";
            //ws.Cells["D1"].Value = "Created";
            //ws.Cells["E1"].Value = "Last modified";
            //ws.Cells["B1:E1"].Style.Font.Bold = true;

            pck.Save();
           

            Spire.License.LicenseProvider.SetLicenseKey("CtOzJs2BlzPokWgBAKMfmNxjRwLa3eqzrAvKtn54UDB/dWjIyGokcs+UQuYuvMY03wX56Ox75KV+U1r5H0PR++c1zc6i8e0QIOVuhMp9Qbg5A9bJJA7e7KvC4KMINTr4jnJy/yTGFwT1aEusw144kml/6oAttwEUoXBkDPLWGOsvNgH1iTYkTGWMXEV8Or4p4t4doNsl0Z7V5qWDKwB6sD/ZiH7l/Jum27FWevOlKIa2VG1rEKjtURYukbWXeSH54IKtmn7nmr0wKwnRgdu3q60aC/PdkxC0zX75EnbU5M6fa3pplU40f3LGOWcgZ2f+8oI7qpPXJ8/s7LrsxBqpQ2YGKfKuqx5ex9ALrXgjnwjcslmXPYun7flHGIkbvBsCjCpo4Ed+M658sZTGATak6gLmftEqhJ1ZZJJKFgXE5qa/TyCY7wIq1ll+z1VNhnSBZUc1RA4TwSBcFKvrZEHlj9o1WFZ1+QqNAcnzh/n+tG48B0wHLCl6D4hroCfWMoaw/23DRxx1WuWqfkazuz2H8ga1RC2XPs83nB7CHPFNs0sT5lsKbfA3P9jgtza5CEhfjAN/3TiwEP/tvnTZY+VABK97veB77h4LEiVMfQXzKfhm9cNW4ft/ofVU2OfqZ8GjtntoZdPxp1bIwTvI98SnQi/H81w19aHwUqNECTeJBjqqHMxdVKVSBAKJL0TM7RyzoOPKS19OfURAxlEgRUqJF/BM8eU0R+UicIM2h36sTuBKO4g3H6woDMlnx0QG0nqthauTB7oK6QFTwk44UQ1kTAu8LeOJwM2xNu5MLsPmoWwDvmIaTuZIW6VUX8C285c9KkrYAf79YKA3e3yxx6SSQdN/jLbtR7MaeGpxRzX0iEbqL9sG1m5USuYVByvVKQ4ntvfCMlLmUN9UCvJ/m63K27Z2dm6fTXIe/g0smYmnvEQ3JQVnldWOi1TKOMK8RbuU5un5mQZ96pLq0Q7g0NLQZh50UMT+OjAzXHPxmXfV6/deHeE8Gbb3ZYJSg7UXW2sty86uXwkj89x5yJTaMNtm6Kh2QQugn/Vd9n8C8QReNewYxjF827FBpMp9yf+vLf2FSyA50wiA9o9luoXYgRmGuUh+g9+KMWgMK5fxQ2h3cHqADzPcwsDhVfG6HuAgt81vH/M5hFLdQztXdvRKVuYOyyTOnQz9K93LZ2EvbeWz0YByRkGxnve+K8UNo3pyNgaPGRQWr5RbeURNJ4PhmM3dB2oMkwE//+s39ccgADdEJS8s35cjRrVEGs8JicRu6mDNqJfdHUNfLmiySMjG/ePwhYkiB2WhJ9AqpY9N7eQ3TBsAMkr34olS6eSNpaE1BjgJsljB27GDnmMAXNZeifyIYpBcqu6H9SLN5pGBF9WHcPVivjdNpMUrKQ==");

            //Create a Workbook instance

            Workbook workbook = new Workbook();

            workbook.LoadFromFile(destFile);

            //workbook.ConverterSetting.SheetFitToPage = true;

            //Save to PDF
            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.SaveToPdf(destFilePDF);
            System.Diagnostics.Process.Start(destFilePDF);
        }

        private void button2_Click(object sender, EventArgs e)
        {
           

        }
    }
}
