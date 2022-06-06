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

    public class Main_Test : BaseWebAiiTest
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
    
        [CodedStep(@"Click 'Span'")]
        public void Main_Test_CodedStep()
        {
            // Click 'Span0'
            ActiveBrowser.Window.SetFocus();
            Pages.DXCOPEN31.Span0.ScrollToVisible(ArtOfTest.WebAii.Core.ScrollToVisibleType.ElementCenterAtWindowCenter);
              
            ab.Application myexcl = new ab.Application();
            var CurrentDirectory = ExecutionContext.Current.DeploymentDirectory;
            Log.WriteLine("Current line " + CurrentDirectory);
            
            ab.Workbook mywrkbk = myexcl.Workbooks.Open(CurrentDirectory + @"\Data\Book1.xlsx");
            ab.Worksheet mySheet = (ab.Worksheet)mywrkbk.Sheets["TransactionCreation"];
            var extractedText = Pages.DXCOPEN31.Span0.InnerText;
            mySheet.Cells[3, 2] = extractedText;
            myexcl.ActiveWorkbook.Save();
            myexcl.ActiveWorkbook.Close();
            myexcl.Application.Quit();
            
        }
    }
}
