using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static Google.Apis.Requests.BatchRequest;
using Newtonsoft.Json;
using System.Linq;
using System.DirectoryServices;
using System.IO;
using System.Reflection;


namespace ExcelAddIn1
{
    public partial class UserControl1 : UserControl
    {

        private AppSettings _appSettings;


        public UserControl1()
        {
            InitializeComponent();
            InitializeSearchControls(); LoadConfiguration();
        }

        private void LoadConfiguration()
        {
            try
            {
                string path = "C:\\Users\\user\\Desktop\\Projects\\C#\\Project\\ExcelAddIn1\\ExcelAddIn1\\appSettings.json";
                string json = File.ReadAllText(path);
                _appSettings = JsonConvert.DeserializeObject<AppSettings>(json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading configuration: {ex.Message}", "Configuration Error");
            }
        }

        //private TextBox 
        private void InitializeSearchControls()
        {
                      
            button.Click += button_Click;
            textBox.Click += textBox_TextChanged;
            Controls.Add(button);

        }


        

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void button_Click(object sender, EventArgs e)
        {

            string searchTerm = textBox.Text;
            
            string apiUrl = _appSettings.ApiUrl;

            if (!string.IsNullOrEmpty(searchTerm))
            {
                try
                {
                    using (HttpClient client = new HttpClient())
                    {

                        HttpResponseMessage response = await client.GetAsync(apiUrl);

                        if (response.IsSuccessStatusCode)
                        {
                            string result = await response.Content.ReadAsStringAsync();

                            var outerResponse = JsonConvert.DeserializeObject<OuterResponse>(result);
                            var innerJson = outerResponse?.searchResponse;

                            var searchResults = JsonConvert.DeserializeObject<SearchResponse>(innerJson)?.items;

                            AddResultToExcel(searchResults);


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

        private void AddResultToExcel(List<SearchResult> searchResults)
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

                activeWorksheet.Columns[1].ColumnWidth = 50;  // Adjust the width as needed

                // Set column width for Link
                activeWorksheet.Columns[2].ColumnWidth = 100;  // Adjust the width as needed


                //ListOfObject ishlatgan holda worksheetga malumotlar qo'shish
                Excel.ListObject table = activeWorksheet.
                    ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    activeWorksheet.UsedRange,
                    Type.Missing, Excel.XlYesNoGuess.xlYes);
                table.Name = "SearchResults";

                //2-qatordan boshlasin
                int rowIndex = 2;


                if (searchResults != null && searchResults.Any())
                {
                    foreach (var result in searchResults)
                    {
                        activeWorksheet.Cells[rowIndex, 1].Value = result.Title;
                        activeWorksheet.Cells[rowIndex, 2].Value = result.Link;

                        rowIndex++;
                    }
                    //MessageBox.Show("Search results added to Excel!", "Success");
                }
                else
                {
                    MessageBox.Show("No search results to add to Excel.", "Information");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding results to Excel: {ex.Message}", "Error");
            }
        }

        public class OuterResponse
        {
            public string Message { get; set; } 
            public string searchResponse { get; set; }
        }

        public class SearchResponse
        {
            public List<SearchResult> items { get; set; }
            // Add other properties as needed
        }

        public class SearchResult
        {
            public string Title { get; set; }
            public string Link { get; set; }
            // Add other properties as needed
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {

        }

       
    }
}

