using EnvDTE;
using Microsoft.VisualStudio.Setup.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace vse
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!args.Any()) return;
            if (args[0] == "run")
            {
                DTE newDte = LaunchVsDte(isPreRelease: false);

                newDte.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateNormal;
                //dte.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateMaximize;
                //dte.MainWindow.WindowState = EnvDTE.vsWindowState.vsWindowStateMinimize;
                //dte.Quit();
                return;
            }
            DTE dte = GetDte();
            if (args[0] == "edit")
            {
                OpenFileAtLine(dte, args[1], int.Parse(args[2]));
            }
            //else if (args[0] == "preview")
            //{
            //    EnvDTE.DTE dte = GetDte();
            //    Console.WriteLine("Previewing file: \"" + args[1] + "\" on line: " + args[2]);
            //    dte.Windows.Item(vsWindowType.vsWindowTypePreview).WindowState = vsWindowState.vsWindowStateNormal;
            //    dte.MainWindow.Activate();
            //    EnvDTE.Window w = dte.ItemOperations.OpenFile("", EnvDTE.Constants.vs);
            //    ((EnvDTE.TextSelection)dte.ActiveDocument.Selection).GotoLine(fileline, true);
            //}
            else if (args[0] == "currentfile")
            {
                Console.WriteLine(dte.ActiveDocument.FullName);
            }
            else if (args[0] == "preview")
            {
                Preview(dte, args[1], int.Parse(args[2]));
            }


        }

        private class PreviewWindow : Window
        {
            public PreviewWindow()
            {
                Type = vsWindowType.vsWindowTypePreview;
                WindowState = vsWindowState.vsWindowStateMaximize;
                Caption = "Test";
            }
            public void SetFocus()
            {
                throw new NotImplementedException();
            }

            public void SetKind(vsWindowType eKind)
            {
                throw new NotImplementedException();
            }

            public void Detach()
            {
                throw new NotImplementedException();
            }

            public void Attach(int lWindowHandle)
            {
                throw new NotImplementedException();
            }

            public void Activate()
            {
            }

            public void Close(vsSaveChanges SaveChanges = vsSaveChanges.vsSaveChangesNo)
            {
                throw new NotImplementedException();
            }

            public void SetSelectionContainer(ref object[] Objects)
            {
                throw new NotImplementedException();
            }

            public void SetTabPicture(object Picture)
            {
                throw new NotImplementedException();
            }

            public Windows Collection { get; }
            public bool Visible { get; set; }
            public int Left { get; set; }
            public int Top { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
            public vsWindowState WindowState { get; set; }
            public vsWindowType Type { get; }
            public LinkedWindows LinkedWindows { get; }
            public Window LinkedWindowFrame { get; }
            public int HWnd { get; }
            public string Kind { get; }
            public string ObjectKind { get; }
            public object Object { get; }
            public ProjectItem ProjectItem { get; }
            public Project Project { get; }
            public DTE DTE { get; }
            public Document Document { get; }
            public object Selection { get; }
            public bool Linkable { get; set; }
            public string Caption { get; set; }
            public bool IsFloating { get; set; }
            public bool AutoHides { get; set; }
            public ContextAttributes ContextAttributes { get; }

            public object get_DocumentData(string bstrWhichData = "")
            {
                throw new NotImplementedException();
            }
        }

        private static void Preview(EnvDTE.DTE dte, string file, int line)
        {
            //Window window = dte.Windows.Item("{9DDABE98-1D02-11D3-89A1-00C04F688DDE}");
            //window.Activate();
            //dte.MainWindow.SetKind(vsWindowType.vsWindowTypePreview);

            Window prev = new PreviewWindow();

            file = $"\"{file}\"";
            Console.WriteLine("Opening file: " + file + " on line: " + line);
            prev.Visible = true;
            prev.WindowState = vsWindowState.vsWindowStateNormal;
            prev.Activate();
            //dte.ExecuteCommand("File.OpenFile", file);
            //dte.ExecuteCommand("Edit.GoTo", line.ToString());
        }

        private static EnvDTE.DTE GetDte()
        {
            //the following line works fine for visual studio 2019:
            EnvDTE.DTE dte = (EnvDTE.DTE)Marshal.GetActiveObject("VisualStudio.DTE.16.0");

            //The number needs to be rolled to the next version each time a new version of visual studio is used... 
            //EnvDTE.DTE dte = null;

            if (dte != null) return dte;

            for (int i = 25; i > 8; i--)
            {
                try
                {
                    dte = (EnvDTE.DTE)Marshal.GetActiveObject("VisualStudio.DTE." + i.ToString() + ".0");
                }
                catch (Exception ex)
                {
                    //don't care... just keep bashing head against wall until success
                }
            }


            return dte;
        }

        private static void OpenFileAtLine(EnvDTE.DTE dte, string file, int line)
        {
            if (dte == null)
            {
                Console.WriteLine("DTE is null");
                return;
            }
            file = $"\"{file}\"";
            Console.WriteLine("Opening file: " + file + " on line: " + line);
            dte.MainWindow.Visible = true;
            dte.ExecuteCommand("File.OpenFile", file);
            dte.ExecuteCommand("Edit.GoTo", line.ToString());
        }


        private static EnvDTE.DTE LaunchVsDte(bool isPreRelease)
        {
            ISetupInstance setupInstance = GetSetupInstance(isPreRelease);
            string installationPath = setupInstance.GetInstallationPath();
            string executablePath = Path.Combine(installationPath, @"Common7\IDE\devenv.exe");
            System.Diagnostics.Process vsProcess = System.Diagnostics.Process.Start(executablePath);
            string runningObjectDisplayName = $"VisualStudio.DTE.16.0:{vsProcess.Id}";

            IEnumerable<string> runningObjectDisplayNames = null;
            object runningObject;
            for (int i = 0; i < 60; i++)
            {
                try
                {
                    runningObject = GetRunningObject(runningObjectDisplayName, out runningObjectDisplayNames);
                }
                catch
                {
                    runningObject = null;
                }

                if (runningObject != null)
                {
                    return (EnvDTE.DTE)runningObject;
                }

                System.Threading.Thread.Sleep(millisecondsTimeout: 1000);
            }

            throw new TimeoutException($"Failed to retrieve DTE object. Current running objects: {string.Join(";", runningObjectDisplayNames)}");
        }

        internal static string GetActiveTextEditor(EnvDTE.DTE dte)
        {
            string docName = dte.ActiveDocument.Name;
            return docName;
        }

        private static object GetRunningObject(string displayName, out IEnumerable<string> runningObjectDisplayNames)
        {
            IBindCtx bindContext = null;
            NativeMethods.CreateBindCtx(0, out bindContext);

            IRunningObjectTable runningObjectTable = null;
            bindContext.GetRunningObjectTable(out runningObjectTable);

            IEnumMoniker monikerEnumerator = null;
            runningObjectTable.EnumRunning(out monikerEnumerator);

            object runningObject = null;
            List<string> runningObjectDisplayNameList = new List<string>();
            IMoniker[] monikers = new IMoniker[1];
            IntPtr numberFetched = IntPtr.Zero;
            while (monikerEnumerator.Next(1, monikers, numberFetched) == 0)
            {
                IMoniker moniker = monikers[0];

                string objectDisplayName = null;
                try
                {
                    moniker.GetDisplayName(bindContext, null, out objectDisplayName);
                }
                catch (UnauthorizedAccessException)
                {
                    // Some ROT objects require elevated permissions.
                }

                if (!string.IsNullOrWhiteSpace(objectDisplayName))
                {
                    runningObjectDisplayNameList.Add(objectDisplayName);
                    if (objectDisplayName.EndsWith(displayName, StringComparison.Ordinal))
                    {
                        runningObjectTable.GetObject(moniker, out runningObject);
                        if (runningObject == null)
                        {
                            throw new InvalidOperationException($"Failed to get running object with display name {displayName}");
                        }
                    }
                }
            }

            runningObjectDisplayNames = runningObjectDisplayNameList;
            return runningObject;
        }

        private static ISetupInstance GetSetupInstance(bool isPreRelease)
        {
            return GetSetupInstances().First(i => IsPreRelease(i) == isPreRelease);
        }

        private static IEnumerable<ISetupInstance> GetSetupInstances()
        {
            ISetupConfiguration setupConfiguration = new SetupConfiguration();
            IEnumSetupInstances enumerator = setupConfiguration.EnumInstances();

            int count;
            do
            {
                ISetupInstance[] setupInstances = new ISetupInstance[1];
                enumerator.Next(1, setupInstances, out count);
                if (count == 1 &&
                    setupInstances != null &&
                    setupInstances.Length == 1 &&
                    setupInstances[0] != null)
                {
                    yield return setupInstances[0];
                }
            }
            while (count == 1);
        }

        private static bool IsPreRelease(ISetupInstance setupInstance)
        {
            ISetupInstanceCatalog setupInstanceCatalog = (ISetupInstanceCatalog)setupInstance;
            return setupInstanceCatalog.IsPrerelease();
        }

        private static class NativeMethods
        {
            [DllImport("ole32.dll")]
            public static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

            [DllImport("ole32.dll")]
            public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
        }
    }
}