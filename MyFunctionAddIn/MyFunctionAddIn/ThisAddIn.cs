using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace MyFunctionAddIn
{
    public partial class ThisAddIn
    {
        MyFunction.MyFunction functionsAddinRef = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            functionsAddinRef = new MyFunction.MyFunction();
            // get the name and GUIF from the class
            string NAME = functionsAddinRef.GetType().Namespace + "." + functionsAddinRef.GetType().Name;
            string GUID = functionsAddinRef.GetType().GUID.ToString().ToUpper();

            // is the add-in already loaded in Excel, but maybe disabled
            // if this is the case - try to re-enable it
            bool fFound = false;
            foreach (Excel.AddIn a in Application.AddIns)
            {
                try
                {
                    if (a.CLSID.Contains(GUID))
                    {
                        fFound = true;
                        if (!a.Installed)
                            a.Installed = true;
                        break;
                    }
                }
                catch { }
            }

            // if we do not see the UDF class in the list of installed addin we need to
            // add it to the collection
            if (!fFound)
            {
                // first register it
                functionsAddinRef.Register();
                // then install it
                Application.AddIns.Add(NAME).Installed = true;
            }
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
