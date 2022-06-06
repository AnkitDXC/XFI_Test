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

namespace TestProject13
{

    public class Policy1 : BaseWebAiiTest
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
            string value = Data["Product"].ToString();
Manager.Desktop.KeyBoard.TypeText(value,150,1);
        }
    
        [CodedStep(@"Desktop command: LeftClick on Text")]
        public void Policy1_CodedStep1()
        {
            // Desktop command: LeftClick on Text
            Pages.DXCOPEN9.Text.Wait.ForExists(30000);
            Pages.DXCOPEN9.Text.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick, 0, 0, ArtOfTest.Common.OffsetReference.AbsoluteCenter);
            
        }
    }
}
