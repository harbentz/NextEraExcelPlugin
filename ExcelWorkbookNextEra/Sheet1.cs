using System;
using System.Data;
using System.Drawing;
using System.Net;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorkbookNextEra
{
    public partial class Sheet1
    {
        NamedRange namedRangeA = null;
        NamedRange namedRangeB = null;
        string[] symbols = new string[] { "MSFT", "BA", "FPL", "NEE" };

        private void UpdateStock()
        {
            try
            {
                WebClient client = new WebClient();
                client.Headers["content-type"] = "application/json";
                for (int i = 0; i < symbols.Length; i++)
                {
                    string symbol = symbols[i];
                    string quoteUrl = $"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey=0GWQD1ERTXN8GVAG";
                    string response = client.DownloadString(quoteUrl);
                    dynamic json = JsonConvert.DeserializeObject(response);
                    try
                    {
                        // https://www.alphavantage.co/documentation/
                        // Maximum number of calls: 5 per minute, 500 per day
                        string quoteStr = json["Global Quote"]["05. price"].Value;
                        namedRangeB.Cells[i + 2].Value2 = decimal.Parse(quoteStr);
                    }
                    catch (Exception) {}
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception: {ex.Message}");
            }
        }

        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            string rangeA = $"A1:A{symbols.Length+1}";
            string rangeB = $"B1:B{symbols.Length+1}";
            namedRangeA = this.Controls.AddNamedRange(this.Range[rangeA], "NamedRangeA");
            namedRangeB = this.Controls.AddNamedRange(this.Range[rangeB], "NamedRangeB");

            namedRangeA.Cells[1].Value2 = "Symbol";
            namedRangeB.Cells[1].Value2 = "Price";
            for (int i = 0; i < symbols.Length; i++)
            {
                namedRangeA.Cells[i+2].Value2 = symbols[i];
                namedRangeB.Cells[i+2].NumberFormat = "$#,##0.00";
            }

            Button button = new Button();
            this.Controls.AddControl(button, 200, 5, 100, 30, "MyButton");
            button.Text = "Update Stock";
            button.Click += new EventHandler(button_Click);
        }

        void button_Click(object sender, EventArgs e)
        {
            UpdateStock();
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet1_Startup);
            this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
        }

        #endregion

    }
}
