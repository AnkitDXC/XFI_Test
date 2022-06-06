using Telerik.TestStudio.Translators.Common;
using Telerik.TestingFramework.Controls.TelerikUI.Blazor;
using Telerik.TestingFramework.Controls.KendoUI.Angular;
using Telerik.TestingFramework.Controls.KendoUI;
using Telerik.WebAii.Controls.Html;
using Telerik.WebAii.Controls.Xaml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Core;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.Design;
using ArtOfTest.WebAii.Design.Execution;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;
using ab=Microsoft.Office.Interop.Excel;
using System.IO;

namespace TestProject13
{

    public class Unique_data : BaseWebAiiTest
    {
        #region [ Dynamic Pages Reference ]

        private Pages _pages;

        /// <summary>
        /// Gets the Pages object that has references
        /// to all the elements, frames or regions
        /// in this project.
        /// </summary>
        public Pages Pages
        {
            get
            {
                if (_pages == null)
                {
                    _pages = new Pages(Manager.Current);
                }
                return _pages;
            }
        }

        #endregion
        
        // Add your test methods here...
    
        [CodedStep(@"New Coded Step")]
        public void Unique_data_CodedStep()
        {
            
            ab.Application myexcl = new ab.Application();
            var CurrentDirectory = ExecutionContext.Current.DeploymentDirectory;

            Log.WriteLine("Current line " + CurrentDirectory);
            
                ab.Workbook mywrkbk = myexcl.Workbooks.Open(CurrentDirectory + @"\Data\Book1.xlsx");
                ab.Worksheet mySheet = (ab.Worksheet)mywrkbk.Sheets["TransactionCreation"];
                ab.Worksheet mySheet2 = (ab.Worksheet)mywrkbk.Sheets["TransactionCreation"];

                ab.Worksheet mySheet1 = (ab.Worksheet)mywrkbk.Sheets["TechnicalTransaction"];
              //  string Var_Name = GetExtractedValue("Policy_Ref").ToString();
                var Var_Name = "Ankit" + Guid.NewGuid();
              //  var varName_2 = "Ankit" + DateTime.Today.ToString("MM/dd/yyyy HH:mm:ss");   
                var varName_2 = "Ankit" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            
                mySheet.Cells[3, 2] = Var_Name;
                mySheet2.Cells[4, 2] = varName_2;
               

                myexcl.ActiveWorkbook.Save();
                myexcl.ActiveWorkbook.Close();
                myexcl.Application.Quit();
            
        }
    
        [CodedStep(@"Desktop command: LeftClick on TxtnameText")]
        public void Unique_data_CodedStep1()
        {
            // Desktop command: LeftClick on TxtnameText
            Pages.FlightTicketsBooking.TxtnameText.Wait.ForExists(30000);
            Pages.FlightTicketsBooking.TxtnameText.ScrollToVisible(ArtOfTest.WebAii.Core.ScrollToVisibleType.ElementCenterAtWindowCenter);
            Pages.FlightTicketsBooking.TxtnameText.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick, 0, 0, ArtOfTest.Common.OffsetReference.AbsoluteCenter);
            
        }
    
        [CodedStep(@"Enter text 'dfdfdf' in 'TxtnameText'")]
        public void Unique_data_CodedStep2()
        {
            // Enter text 'dfdfdf' in 'TxtnameText'
            Actions.SetText(Pages.FlightTicketsBooking.TxtnameText, "");
            Pages.FlightTicketsBooking.TxtnameText.ScrollToVisible(ArtOfTest.WebAii.Core.ScrollToVisibleType.ElementCenterAtWindowCenter);
            ActiveBrowser.Window.SetFocus();
            Pages.FlightTicketsBooking.TxtnameText.Focus();
            Pages.FlightTicketsBooking.TxtnameText.MouseClick();
            Manager.Desktop.KeyBoard.TypeText("dfdfdf", 50, 100, true);
            
        }
    }
}
