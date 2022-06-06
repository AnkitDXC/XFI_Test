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

namespace TestProject13
{

    public class Login_Policy : BaseWebAiiTest
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
    
        [CodedStep(@"Enter text '' in 'Text' - DataDriven: [$(Product)]")]
        public void Policy1_CodedStep()
        {
                        // Enter text '' in 'Text'
            //            Actions.SetText(Pages.DXCOPEN9.Text, "");
            //            Pages.DXCOPEN9.Text.ScrollToVisible(ArtOfTest.WebAii.Core.ScrollToVisibleType.ElementCenterAtWindowCenter);
            //            ActiveBrowser.Window.SetFocus();
            //            Pages.DXCOPEN9.Text.Focus();
            //            Pages.DXCOPEN9.Text.MouseClick();
            //            Manager.Desktop.KeyBoard.TypeText(((string)(System.Convert.ChangeType(Data["Product"], typeof(string)))), 50, 100, true);
                        //string value = Data["Product"].ToString();
          //  Manager.Desktop.KeyBoard.TypeText(value,150,1);
            
            
            
                ab.Application myexcl = new ab.Application();
                ab.Workbook mywrkbk = myexcl.Workbooks.Open(@"C:\Users\a225\OneDrive - DXC Production\Documents\Book1.xlsx");
                ab.Worksheet mySheet = (ab.Worksheet)mywrkbk.Sheets["TransactionCreation"];
                ab.Worksheet mySheet2 = (ab.Worksheet)mywrkbk.Sheets["TransactionCreation"];

                ab.Worksheet mySheet1 = (ab.Worksheet)mywrkbk.Sheets["TechnicalTransaction"];
              //  string Var_Name = GetExtractedValue("Policy_Ref").ToString();
                string Var_Name = "Ankit";
            
                mySheet.Cells[3, 2] = Var_Name;
                mySheet2.Cells[4, 2] = "xyz";
                  



                myexcl.ActiveWorkbook.Save();
                myexcl.ActiveWorkbook.Close();
                myexcl.Application.Quit();
            
        }
    }
}
