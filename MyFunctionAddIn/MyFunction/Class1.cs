using Microsoft.Win32;
using System;
using Extensibility;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MyFunction
{
    [Guid("78A7C4F1-C5D6-4C86-AADE-696C2616679D")]
    public interface IFunctions
    {
        double SharePrice(string companyCode, string priceDate);
    }

    // In memory list of prices
    public static class EODPriceCache
    {
        public static List<Price> EODPrices = new List<Price>
    {
        new Price() {CompanyCode = "DEMO", PriceDate = "06/12/2020", EODPrice = 114.89d},
        new Price() {CompanyCode = "DEMO", PriceDate = "06/13/2020", EODPrice = 107.89d},
        new Price() {CompanyCode = "DEMO", PriceDate = "06/14/2020", EODPrice = 119.89d},
        new Price() {CompanyCode = "DEMO", PriceDate = "06/15/2020", EODPrice = 110.89d},
        new Price() {CompanyCode = "NXT", PriceDate = "06/12/2020", EODPrice = 5480},
        new Price() {CompanyCode = "NXT", PriceDate = "06/13/2020", EODPrice = 5000},
        new Price() {CompanyCode = "NXT", PriceDate = "06/14/2020", EODPrice = 5400},
        new Price() {CompanyCode = "NXT", PriceDate = "06/15/2020", EODPrice = 5480},
        new Price() {CompanyCode = "TSCO", PriceDate = "06/12/2020", EODPrice = 200.0d},
        new Price() {CompanyCode = "TSCO", PriceDate = "06/13/2020", EODPrice = 237.2d},
        new Price() {CompanyCode = "TSCO", PriceDate = "06/14/2020", EODPrice = 217.7d},
        new Price() {CompanyCode = "TSCO", PriceDate = "06/15/2020", EODPrice = 227.7d},
        new Price() {CompanyCode = "MSFT", PriceDate = "06/15/2020", EODPrice = 2200d},
        new Price() {CompanyCode = "MSFT", PriceDate = "06/15/2020", EODPrice = 2234d},
        new Price() {CompanyCode = "MSFT", PriceDate = "06/15/2020", EODPrice = 2250d},
        new Price() {CompanyCode = "MSFT", PriceDate = "06/15/2020", EODPrice = 2600d},
    };
    }
    //=MyFunction.MyFunction.SharePrice("DEMO", "06/15/2020")
    //=SharePrice("DEMO", "06/15/2020")


    [Guid("D1205017-C098-48FD-BD4F-ABCB739F7F19"),
    ProgId("MyFunction.SharePrice"),
    ClassInterface(ClassInterfaceType.AutoDual),
    ComVisible(true)]
    public class MyFunction: IFunctions, Extensibility.IDTExtensibility2
    {
        public MyFunction() { }

        public double SharePrice(string companyCode, string priceDate)
        {// This could be threaded if not performant enough
            Price selectedEOD;
            selectedEOD = EODPriceCache.EODPrices.Find(c => (c.CompanyCode == companyCode) && (c.PriceDate == priceDate));
            return selectedEOD.EODPrice;
        }

        #region IDTExtensibility2
        private static Excel.Application Application; // ref to Excel
        private static object ThisAddIn;
        private static bool fVstoRegister = false;

        /// <summary>
        /// When we finally do connect and load in Excel we want to get the
        /// reference to the application, so that we can use the application
        /// instace in our UDF as needed
        /// </summary>
        /// <param name="application"></param>
        /// <param name="connectMode"></param>
        /// <param name="addInInst"></param>
        /// <param name="custom"></param>
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            // get a reference to the instance of the add-in
            Application = application as Excel.Application;
            ThisAddIn = addInInst;
        }       

        /// <summary>
        /// We call this from VSTO so that we can ge the DLL
        /// to register itself and load every time
        /// </summary>
        public void Register() // exposed to VSTO
        {
            fVstoRegister = true;
            RegisterFunction(typeof(MyFunction));
        }

        /// <summary>
        /// When we disconnect - remove everything - clean up
        /// </summary>
        /// <param name="disconnectMode"></param>
        /// <param name="custom"></param>
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            // clean up
            Marshal.ReleaseComObject(Application);
            Application = null;
            ThisAddIn = null;
            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public void OnAddInsUpdate(ref System.Array custom) { }
        public void OnStartupComplete(ref System.Array custom) { }
        public void OnBeginShutdown(ref System.Array custom) { }


        /// <summary>
		/// Registers the COM Automation Add-in in the CURRENT USER context
		/// and then registers it in all versions of Excel on the users system
		/// without the need of administrator permissions
		/// </summary>
		/// <param name="type"></param>
		[ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            string PATH = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.Replace("\\", "/");
            string ASSM = Assembly.GetExecutingAssembly().FullName;
            int startPos = ASSM.ToLower().IndexOf("version=") + "version=".Length;
            int len = ASSM.ToLower().IndexOf(",", startPos) - startPos;
            string VER = ASSM.Substring(startPos, len);
            string GUID = "{" + type.GUID.ToString().ToUpper() + "}";
            string NAME = type.Namespace + "." + type.Name;
            string BASE = @"Classes\" + NAME;
            string CLSID = @"Classes\CLSID\" + GUID;

            // open the key
            RegistryKey CU = Registry.CurrentUser.OpenSubKey("Software", true);

            // is this version registred?
            RegistryKey key = CU.OpenSubKey(CLSID + @"\InprocServer32\" + VER);
            if (key == null)
            {
                // The version of this class currently being registered DOES NOT
                // exist in the registry - so we will now register it

                // BASE KEY
                // HKEY_CURRENT_USER\CLASSES\{NAME}
                key = CU.CreateSubKey(BASE);
                key.SetValue("", NAME);

                // HKEY_CURRENT_USER\CLASSES\{NAME}\CLSID}
                key = CU.CreateSubKey(BASE + @"\CLSID");
                key.SetValue("", GUID);

                // CLSID
                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}
                key = CU.CreateSubKey(CLSID);
                key.SetValue("", NAME);

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\Implemented Categories
                key = CU.CreateSubKey(CLSID + @"\Implemented Categories").CreateSubKey("{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}");

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\InProcServer32
                key = CU.CreateSubKey(CLSID + @"\InprocServer32");
                key.SetValue("", @"c:\Windows\SysWow64\mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
                key.SetValue("Class", NAME);
                key.SetValue("CodeBase", PATH);
                key.SetValue("Assembly", ASSM);
                key.SetValue("RuntimeVersion", "v4.0.30319");

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\InProcServer32\{VERSION}
                key = CU.CreateSubKey(CLSID + @"\InprocServer32\" + VER);
                key.SetValue("Class", NAME);
                key.SetValue("CodeBase", PATH);
                key.SetValue("Assembly", ASSM);
                key.SetValue("RuntimeVersion", "v4.0.30319");

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\ProgId
                key = CU.CreateSubKey(CLSID + @"\ProgId");
                key.SetValue("", NAME);

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\Progammable
                key = CU.CreateSubKey(CLSID + @"\Programmable");

                // now register the addin in the addins sub keys for each version of Office
                foreach (string keyName in Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\").GetSubKeyNames())
                {
                    if (IsVersionNum(keyName))
                    {
                        // if the adding i found in the Add-in Manager - remove it
                        key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + keyName + @"\Excel\Add-in Manager", true);
                        if (key != null)
                        {
                            key.SetValue(NAME, "");
                        }
                    }
                }
                if (!fVstoRegister)
                {
                    // all done - this just helps to assure REGASM is complete
                    // this is not needed, but is useful for troubleshooting
                    MessageBox.Show("Registered " + NAME + ".");
                }
            }
        }


        /// <summary>
        /// Unregisters the add-in, by removing all the keys
        /// </summary>
        /// <param name="type"></param>
        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            string GUID = "{" + type.GUID.ToString().ToUpper() + "}";
            string NAME = type.Namespace + "." + type.Name;
            string BASE = @"Classes\" + NAME;
            string CLSID = @"Classes\CLSID\" + GUID;
            // open the key
            RegistryKey CU = Registry.CurrentUser.OpenSubKey("Software", true);
            // DELETE BASE KEY
            // HKEY_CURRENT_USER\CLASSES\{NAME}
            try
            {
                CU.DeleteSubKeyTree(BASE);
            }
            catch { }
            // HKEY_CURRENT_USER\CLASSES\{NAME}\CLSID}
            try
            {
                CU.DeleteSubKeyTree(CLSID);
            }
            catch { }
            // now un-register the addin in the addins sub keys for Office
            // here we just make sure to remove it from allversions of Office
            foreach (string keyName in Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\").GetSubKeyNames())
            {
                if (IsVersionNum(keyName))
                {
                    RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + keyName + @"\Excel\Add-in Manager", true);
                    if (key != null)
                    {
                        try
                        {
                            key.DeleteValue(NAME);
                        }
                        catch { }
                    }
                    key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + keyName + @"\Excel\Options", true);
                    if (key == null)
                        continue;
                    foreach (string valueName in key.GetValueNames())
                    {
                        if (valueName.StartsWith("OPEN"))
                        {
                            if (key.GetValue(valueName).ToString().Contains(NAME))
                            {
                                try
                                {
                                    key.DeleteValue(valueName);
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            MessageBox.Show("Unregistered " + NAME + "!");
        }

        /// <summary>
        /// HELPER FUNCTION
        /// This assists is in determining if the subkey string we are passed
        /// is of the type like:
        ///     8.0
        ///     11.0
        ///     14.0
        ///     15.0
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool IsVersionNum(string s)
        {
            int idx = s.IndexOf(".");
            if (idx >= 0 && s.EndsWith("0") && int.Parse(s.Substring(0, idx)) > 0)
                return true;
            else
                return false;
        }

        #endregion

    }
}
