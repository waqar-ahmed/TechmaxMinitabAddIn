using System;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Reflection;
using RGiesecke.DllExport;
using MinitabAddinTLB;
using Mtb;


namespace Techmax
{
    public static class Globals
    {
        public const string GUID = "868FB674-34B0-47D6-A8AA-F42A16975857";
        public const string Namespace = "Techmax";
        public const string NamespaceAndClass = "Techmax.AddIn";
    }

    [type: ComVisible(true)]
    [type: GuidAttribute(Globals.GUID)]
    [type: ClassInterface(ClassInterfaceType.None)]
    [type: ProgId(Globals.NamespaceAndClass)]

    public class AddIn : IMinitabAddin
    {
        internal static Mtb.Application gMtbApp;

        [method: DllExportAttribute("DllRegisterServer", CallingConvention = CallingConvention.StdCall)]
        public static Int32 DllRegisterServer()
        {
            try
            {
                RegistryKey key = Registry.ClassesRoot.CreateSubKey(Globals.NamespaceAndClass);
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.ClassesRoot.CreateSubKey(Globals.NamespaceAndClass + @"\CLSID");
                key.SetValue("", "{" + Globals.GUID + "}");

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + "}");
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories");

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}");

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32");
                key.SetValue("", "mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32\1.0.0.0");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.ClassesRoot.CreateSubKey(@"Wow6432Node\CLSID\{" + Globals.GUID + @"}\ProdId");
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.LocalMachine.CreateSubKey(@"\SOFTWARE\Classes\" + Globals.NamespaceAndClass);
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.LocalMachine.CreateSubKey(@"\SOFTWARE\Classes\" + Globals.NamespaceAndClass + @"\CLSID");
                key.SetValue("", "{" + Globals.GUID + "}");

                key = Registry.LocalMachine.CreateSubKey(@"\SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + "}");
                key.SetValue("", Globals.NamespaceAndClass);

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories");

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}");

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32");
                key.SetValue("", "mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\InprocServer32\1.0.0.0");
                key.SetValue("Class", Globals.NamespaceAndClass);
                key.SetValue("Assembly", Globals.Namespace + ", Culture=neutral, PublicKeyToken=null");
                key.SetValue("RuntimeVersion", Environment.Version.ToString());
                key.SetValue("CodeBase", @"file:///" + Assembly.GetExecutingAssembly().Location);

                key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Classes\Wow6432Node\CLSID\{" + Globals.GUID + @"}\ProgId");
                key.SetValue("", Globals.NamespaceAndClass);
            }
            catch (Exception ex)
            {
                // Probably didn't have permissions to modify the registry
            } 
        
            return 0;
        }

        public AddIn()
        {
        }

        ~AddIn()
        {
        }

        public void OnConnect(Int32 iHwnd, Object pApp, ref Int32 iFlags)
        {
            // This method is called as Minitab is initializing your add-in.
            // The “iHwnd” parameter is the handle to the main Minitab window.
            // The “pApp” parameter is a reference to the “Minitab Automation object.”
            // You can hold onto either of these for use in your add-in.
            // “iFlags” is used to tell Minitab if your add-in has dynamic menus (i.e. should be reloaded each time
            // Minitab starts up).  Set Flags to 1 for dynamic menus and 0 for static.
            AddIn.gMtbApp = pApp as Mtb.Application;
            // This forces Minitab to retain all commands (even those run by the interactive user):
            AddIn.gMtbApp.Options.SaveCommands = true;
            // Static menus:
            iFlags = 0;
            return;
        }

        public void OnDisconnect()
        {
            // This method is called as Minitab is closing your add-in.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            try
            {
                Marshal.FinalReleaseComObject(AddIn.gMtbApp);
                AddIn.gMtbApp = null;
            }
            catch
            {
            }
            return;
        }

        public String GetName()
        {
            // This method returns the friendly name of your add-in:
            // Both the name and the description of the add-in are stored in the registry.
            return "Example C♯ Minitab Add-In";
        }

        public String GetDescription()
        {
            // This method returns the description of your add-in:
            return "An example Minitab add-in written in C♯ using the “My Menu” functionality.";
        }

        public void GetMenuItems(ref String sMainMenu, ref Array saMenuItems, ref Int32 iFlags)
        {
            // This method returns the text for the main menu and each menu item.
            // You can return "|" to create a menu separator in your menu items.
            sMainMenu = "&Techmax Menu";  // This string is the name of the menu.

            saMenuItems = new String[5];  // The strings in this array are the names of the items on the aforementioned menu.

            saMenuItems.SetValue("Describe &column(s)…", 0);
            saMenuItems.SetValue("Rename active &worksheet…", 1);
            saMenuItems.SetValue("|", 2);
            saMenuItems.SetValue("&DOS window", 3);
            saMenuItems.SetValue("&Geometric Mean and Mean Absolute Difference…", 4);

            // Flags is not currently used:
            iFlags = 0;

            return;
        }

        public String OnDispatchCommand(Int32 iMenu)
        {
            // This method is called whenever a user selects one of your menu items.
            // The iMenu variable should be equivalent to the menu item index set in “GetMenuItems.”
            String command = String.Empty;
            
            DialogResult dialogResult = new DialogResult();
            switch (iMenu)
            {
                case 0:
                    // Describe column(s):
                    FormDescribe formDescribe = new FormDescribe(ref AddIn.gMtbApp);
                    // Fill up list box in dialog with numeric columns in worksheet:
                    formDescribe.checkedListBoxOfColumns.ClearSelected();
                    Int32 lColumnCount = AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Count;
                    for (Int32 i = 1; i <= lColumnCount; i += 1)
                    {
                        // Select only the numeric columns:
                        if (AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i).DataType == MtbDataTypes.Numeric)
                        {
                            formDescribe.checkedListBoxOfColumns.Items.Add(gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i).SynthesizedName);
                        }
                    }
                    // Show the dialog:
                    dialogResult = formDescribe.ShowDialog();
                    if (dialogResult == DialogResult.OK)
                    {
                        StringBuilder cmnd = new StringBuilder("Describe ");

                        Boolean bPrev = false;
                        for (Int32 i = 0; i < formDescribe.checkedListBoxOfColumns.CheckedItems.Count; i += 1)
                        {
                            if (bPrev)
                            {
                                cmnd.Append(" ");
                            }
                            cmnd.Append(formDescribe.checkedListBoxOfColumns.CheckedItems[i].ToString());
                            bPrev = true;
                        }
                        if (formDescribe.chkMean.Checked)
                        {
                            cmnd.Append("; Mean");
                        }
                        if (formDescribe.chkVariance.Checked)
                        {
                            cmnd.Append("; Variance");
                        }
                        if (formDescribe.chkSum.Checked)
                        {
                            cmnd.Append("; Sums");
                        }
                        if (formDescribe.chkNnonmissing.Checked)
                        {
                            cmnd.Append("; N");
                        }
                        if (formDescribe.chkHistogram.Checked)
                        {
                            cmnd.Append("; GHist");
                        }
                        if (formDescribe.chkBoxplot.Checked)
                        {
                            cmnd.Append("; GBoxplot");
                        }
                        cmnd.Append(".");
                        command = cmnd.ToString();
                    }
                    formDescribe.Close();
                    break;
                case 1:
                    // Rename active worksheet:
                    FormRename formRename = new FormRename(ref AddIn.gMtbApp);
                    String sCurrent = AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Name;
                    formRename.textBoxCurrent.Enabled = true;
                    formRename.textBoxCurrent.Text = sCurrent;
                    formRename.textBoxCurrent.Enabled = false;
                    // Show the dialog:
                    dialogResult = formRename.ShowDialog();
                    if (dialogResult == DialogResult.OK)
                    {
                        AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Name = formRename.textBoxNew.Text;
                    }
                    formRename.Close();
                    break;
                case 2:
                    break;
                case 3:
                    // Open a DOS Window:
                    String[] fileNamePossibilities = { "cmd.exe", "command.com" };
                    Process process;
                    ProcessStartInfo processStartInfo;
                    foreach (String fileNamePossibility in fileNamePossibilities)
                    {
                        process = new Process();
                        processStartInfo = new ProcessStartInfo();
                        processStartInfo.UseShellExecute = true;
                        processStartInfo.FileName = fileNamePossibility;
                        process.StartInfo = processStartInfo;
                        try
                        {
                            process.Start();
                            break;
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(e.Message, "My Menu");
                            MessageBox.Show("Cannot locate DOS executable or otherwise start a command prompt…", "My Menu");
                            continue;
                        }
                    }
                    break;
                case 4:
                    // “Geometric Mean” and “Mean Absolute Difference” (stored in the worksheet):
                    FormGeoMean formGeoMean = new FormGeoMean(ref AddIn.gMtbApp);
                    // Fill up list box in dialog with numeric columns in worksheet:
                    lColumnCount = AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Count;
                    Hashtable hashtableOfNumericColumns = new Hashtable();
                    for (Int32 i = 1; i <= lColumnCount; i += 1)
                    {
                        if (AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i).DataType == MtbDataTypes.Numeric)
                        {
                            String sSynthesizedColumnName = AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i).SynthesizedName;
                            String sColumnName = AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i).Name;
                            // Add column name (if it exists):
                            if (sColumnName != sSynthesizedColumnName)
                            {
                                sSynthesizedColumnName += String.Concat("  ", sColumnName);
                            }
                            formGeoMean.comboBox.Items.Add(sSynthesizedColumnName);
                            hashtableOfNumericColumns.Add(sSynthesizedColumnName, AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i));
                        }
                    }
                    // Show the dialog:
                    dialogResult = formGeoMean.ShowDialog();
                    if (dialogResult == DialogResult.OK)
                    {
                        // Get data from the column and pass it to the function to do calculations:
                        Object selectedItem = formGeoMean.comboBox.SelectedItem;
                        Mtb.Column mtbDataColumn = (Mtb.Column)hashtableOfNumericColumns[selectedItem];
                        // “FindGeoMean” takes an array of doubles and returns the geometric mean.
                        // “bSuccess” indicates if the calculations were completed.
                        Array daData = mtbDataColumn.GetData();
                        Boolean bSuccess;
                        Object dGeoMean = FindGeoMean(ref daData, out bSuccess);
                        if (bSuccess)
                        {
                            // Find the “Mean Absolute Difference”:
                            Object dMAD = FindMAD(ref daData);
                            // Store both values in the first available column:
                            Mtb.Column mtbStorageColumn = AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Add();
                            mtbStorageColumn.SetData(ref dGeoMean, 1, 1);
                            mtbStorageColumn.SetData(ref dMAD, 2, 1);
                            mtbStorageColumn.Name = "MyResults";
                        }
                        else
                        {
                            // An error occurred:
                            AddIn.gMtbApp.ActiveProject.ExecuteCommand("NOTE ** Error ** Cannot compute statistics…");
                        }
                        formGeoMean.Close();
                    }
                    break;
                default:
                    break;
            }
            
            return command;
        }

        public void OnNotify(AddinNotifyType eAddinNotifyType)
        {
            // This method is called when Minitab notifies your add-in that something has changed.
            // Use the “eAddinNotifyType” parameter to figure out what changed.
            // Minitab currently fires no events, so this method is not called.
            return;
        }

        public Boolean QueryCustomCommand(String sCommand)
        {
            // This method is called when Minitab asks your Addin if it supports a custom command.
            // The argument “sCommand” is the name of the custom command.  Return “true” if you support the command.
            return sCommand.ToUpper() == "EXPLORER" || sCommand.ToUpper() == "CLEAR";
        }

        public void ExecuteCustomCommand(String sCommand, ref Array saArgs)
        {
            // This method is called when Minitab asks your add-in to execute a custom command.
            // The argument “sCommand” is the name of the command, and “saArgs” is an array of arguments.
            if (sCommand.ToUpper() == "EXPLORER")
            {
                // Open Windows Explorer:
                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.UseShellExecute = true;
                processStartInfo.FileName = "explorer.exe";
                process.StartInfo = processStartInfo;
                try
                {
                    process.Start();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "My Menu");
                    MessageBox.Show("Apparently, Windows Explorer could not be started…", "My Menu");
                }
            }
            else if (sCommand.ToUpper() == "CLEAR")
            {
                // Clear indicated columns:
                Int32 lColumnCount = AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Count;
                Int32 saArgsCardinality = saArgs.GetLength(saArgs.Rank - 1);
                IEnumerator myEnumerator = saArgs.GetEnumerator();
                while (myEnumerator.MoveNext())
                {
                    for (Int32 i = 1; i <= lColumnCount; i++)
                    {
                        Int32 myEnumeratorCurrent = 0;
                        Int32.TryParse(myEnumerator.Current.ToString(), out myEnumeratorCurrent);
                        if (AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i).Number == myEnumeratorCurrent)
                        {
                            AddIn.gMtbApp.ActiveProject.ActiveWorksheet.Columns.Item(i).Clear();
                        }
                    }
                }
            }
        }

        public Double FindGeoMean(ref Array saData, out Boolean bSuccess)
        {
            // Find geometric mean:
            Double dSum = 0.0;
            Int32 iCount = 0;
            bSuccess = true;
            foreach (Double dValue in saData)
            {
                if (dValue <= 0)
                {
                    bSuccess = false;
                    MessageBox.Show("All values must be strictly positive!", "My Menu");
                    break;
                }
                dSum += Math.Log(dValue);
                iCount += 1;
            }

            return Math.Exp(dSum / iCount);
        }

        public Double FindMAD(ref Array daData)
        {
            // Find M(ean) A(bsolute) D(ifference):
            Double dSum = 0.0;
            Int32 iCount = 0;

            foreach (Double dValue in daData)
            {
                dSum += dValue;
                iCount += 1;
            }

            Double dMAD = 0.0;
            Double dMean = dSum / iCount;

            foreach (Double dValue in daData)
            {
                dMAD += Math.Abs(dValue - dMean);
            }

            dMAD /= iCount;

            return dMAD;
        }
    }
}

