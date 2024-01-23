using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn1
{
    public partial class UserControl1 : UserControl
    {


        public UserControl1()
        {
            InitializeComponent();
            InitializeSearchControls();
        }

        //private TextBox 
        private void InitializeSearchControls()
        {
                      
            button.Click += button1_Click;
            textBox.Click += textBox_TextChanged;


            Controls.Add(button);
            //Controls.Add(textBoxSearch);

        }

        private string ApiKey = "AIzaSyCJliVyqGoNRpVN_5hJ491EG5QUPwoAua4";
        private string SearchEngineId = "7441ce6c6fb45452a";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            // Use a custom form to get the search term from the user
            //string searchTerm = GetUserInput("Enter your search term:");

            string searchTerm = textBox.Text;

            if (!string.IsNullOrEmpty(searchTerm))
            {
                try
                {
                    using (HttpClient client = new HttpClient())
                    {
                        string apiUrl = $"https://www.googleapis.com/customsearch/v1?q={searchTerm}&key={ApiKey}&cx={SearchEngineId}";

                        HttpResponseMessage response = await client.GetAsync(apiUrl);

                        if (response.IsSuccessStatusCode)
                        {
                            string result = await response.Content.ReadAsStringAsync();
                            // Process the JSON result as needed
                            AddResultToExcel(result);

                        }
                        else
                        {
                            MessageBox.Show($"Error: {response.StatusCode.ToString()} - {response.ReasonPhrase}", "Search Error");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}", "Search Error");
                }
            }
        }

        private string GetUserInput(string prompt)
        {
            // Use a custom form to get user input
            using (var form = new Form())
            {
                form.Text = "User Input";
                var label = new System.Windows.Forms.Label() { Left = 20, Top = 20, Text = prompt };
                var textBox = new System.Windows.Forms.TextBox() { Left = 20, Top = 50, Width = 100 };
                var button = new System.Windows.Forms.Button() { Text = "OK", Left = 150, Top = 50 };

                button.Click += (s, e) => { form.Close(); };

                form.Controls.Add(label);
                form.Controls.Add(textBox);
                form.Controls.Add(button);

                form.ShowDialog();

                return textBox.Text;
            }
        }


        private void AddResultToExcel(string jsonResult)
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;

                if (excelApp == null)
                {
                    MessageBox.Show("Excel application not availabe.");
                    return;
                }
                Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

                if (activeWorkbook == null)
                {
                    MessageBox.Show("No active workbook.", "Error");
                    return;
                }
                Excel.Worksheet activeWorksheet = activeWorkbook.ActiveSheet;

                //Malumotlarni ushlab turish uchun List
                List<SearchResult> searchResults = ParseJsonResult(jsonResult);

                //agar yangi worksheet kerak bo'lsa och
                if (activeWorksheet == null)
                {
                    activeWorksheet = activeWorkbook.Worksheets.Add();
                }

                //avvalgi datalarni o'chir
                activeWorksheet.Cells.Clear();

                //har bir qatorga nom
                activeWorksheet.Cells[1, 1].Value = "Title";
                activeWorksheet.Cells[1, 2].Value = "Link";

                //ListOfObject ishlatgan holda worksheetga malumotlar qo'shish
                Excel.ListObject table = activeWorksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, activeWorksheet.UsedRange, Type.Missing, Excel.XlYesNoGuess.xlYes);
                table.Name = "SearchResults";

                //2-qatordan boshlasin
                int rowIndex = 2;


                foreach (var result in searchResults)
                {
                    activeWorksheet.Cells[rowIndex, 1].Value = result.Title;
                    activeWorksheet.Cells[rowIndex, 2].Value = result.Link;

                    rowIndex++;
                }
                MessageBox.Show("Search results added to Excel!", "Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding results to Excel: {ex.Message}", "Error");
            }
        }

        private List<SearchResult> ParseJsonResult(string jsonResult)
        {
            List<SearchResult> results = new List<SearchResult>();

            JObject json = JObject.Parse(jsonResult);
            JArray items = (JArray)json["items"];

            if (items != null)
            {
                foreach (var item in items)
                {
                    string title = (string)item["title"];
                    string link = (string)item["link"];

                    results.Add(new SearchResult { Title = title, Link = link });
                }
            }

            return results;
        }

        private class SearchResult
        {
            public string Title { get; set; }
            public string Link { get; set; }
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void button_Click(object sender, EventArgs e)
        {

        }
    }
}

