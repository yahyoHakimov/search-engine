
using Microsoft.Office.Core;
using System.Data.SqlClient;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane myTaskPane;

       
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           UserControl1 myUserControl  = new UserControl1();
            myTaskPane = (Microsoft.Office.Tools.CustomTaskPane)CustomTaskPanes.Add(myUserControl, "Qidiruv Oynasi");

            myTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            myTaskPane.Width = 400;

            // Make the task pane visible
            myTaskPane.Visible = true;

            //InitializeDatabase();

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
