
using Microsoft.Office.Core;
using System.Data.SqlClient;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane myTaskPane;

        private void InitializeDatabase()
        {
            string connectionString =
                "Server = YAXYOBEK-HAKIMO\\YAHYOSERVER;" +
                " Database = EMPLOYER; Integrated Security = True;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                //user Table
                using (SqlCommand command = new SqlCommand(
                "UPDATE TABLE Employee (UserId INT PRIMARY KEY, Username NVARCHAR(255));", connection))
                {
                    command.ExecuteNonQuery();
                }
                //LogData table
                using (SqlCommand logCommand = new SqlCommand(
                   "UPDATE TABLE LogData (LogId INT PRIMARY KEY, UserId INT, ActivityType NVARCHAR(255), " +
                    "SearchTerm NVARCHAR(255), Timestamp DATETIME, FOREIGN KEY(UserId) REFERENCES Users(UserId));", connection))
                {
                    logCommand.ExecuteNonQuery();
                }

            }


        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           UserControl1 myUserControl  = new UserControl1();
            myTaskPane = (Microsoft.Office.Tools.CustomTaskPane)CustomTaskPanes.Add(myUserControl, "Qidiruv Oynasi");

            myTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            myTaskPane.Width = 400;

            // Make the task pane visible
            myTaskPane.Visible = true;

            InitializeDatabase();

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        
        #endregion
    }
}
