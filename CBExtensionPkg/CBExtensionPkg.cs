//
//
// Project:  CBExtensionPkg
//
using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.Win32;
using Microsoft.VisualStudio.TextManager.Interop;

// ReSharper disable InconsistentNaming

namespace BlackIceSoftware.CBExtensionPkg
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    // https://msdn.microsoft.com/en-us/library/microsoft.visualstudio.shell.interop.uicontextguids80.aspx
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)] // UIContext_SolutionExists
    // [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")] // UIContext_SolutionExists
    [ProvideAutoLoad(UIContextGuids80.CodeWindow)] // UICONTEXT_CodeWindow
    //[ProvideAutoLoad("8fe2df1d-e0da-4ebe-9d5c-415d40e487b5")] // UICONTEXT_CodeWindow
    [ProvideLoadKey("Standard", "1.12", "CBExtensionPkg", "BlackIce Software", 104)]
    [ProvideOptionPage(typeof(AutoSaveOptions), "CB AutoSave", "General", 101, 106, true)]
    [Guid(GuidList.guidCBExtensionPkgPkgString)]
    public sealed class CBExtensionPkg : Package, IVsBroadcastMessageEvents, IOleCommandTarget
    {
        private const string SaveSingleDocumentSubKeyName = "SaveSingleDocument";
        private const string SaveFilesSubKeyName = "SaveFiles";
        private const string SaveProjectsSubKeyName = "SaveProjects";
        private const string SaveSolutionsSubKeyName = "SaveSolutions";
        private const string LostFocusCmdSubKeyName = "LostFocusCmd";
        private const string LostFocusCmdArgsSubKeyName = "LostFocusCmdArgs";
        private const string LostDocFocusCmdSubKeyName = "LostDocFocusCmd";
        private const string LostDocFocusCmdArgsSubKeyName = "LostDocFocusCmdArgs";
        private const string CBVSAddInSubKeyName = "CBVSAddIns";

        //private IVsAdaptor _vsAdaptor;
        //private IVim _vim;
        //IComponentModel_componentModel;

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require
        /// any Visual Studio service because at this point the package object is created but
        /// not sited yet inside Visual Studio environment. The place to do all the other
        /// initialization is the Initialize method.
        /// </summary>
        public CBExtensionPkg()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", ToString()));
            // _componentModel = (IComponentModel)GetService(typeof(SComponentModel));

            //_exportProvider = _componentModel.DefaultExportProvider;
            //_vim = _exportProvider.GetExportedValue<IVim>();
            //_vsAdapter = _exportProvider.GetExportedValue<IVsAdapter>();
        }

        #region Package Members

        private const int WmActivateapp = 0x001C;

        private uint abmCookie;
        private DTE2 _dte;
        private Events events;
        private IVsOutputWindowPane mOutputWindowPane;
        private Guid CBPkgPaneGuid;
        private WindowEvents mWinEvents;
        private IVsShell vsShell;
        private IVsTextManager _textManager;
        private IComponentModel _componentModel;
        private System.ComponentModel.Composition.Hosting.ExportProvider _exportProvider;

        protected override int QueryClose(out bool canClose)
        {
            var opts = GetDialogPage(typeof(AutoSaveOptions)) as AutoSaveOptions;
            Debug.Assert(null != opts, "OptionsPage obj is null");
            if (opts == null) { return base.QueryClose(out canClose); }

            bool autosaveEnabled = opts.AutoSaveDocuments;
            bool autosaveSingleDocEnabled = opts.AutoSaveSingleDocument;
            bool autosaveProjectEnabled = opts.AutoSaveProjects;
            bool autosaveSolutionEnabled = opts.AutoSaveSolution;

#if USE_LOST_FOCUS_CMDS
            string lostFocusCommand = opts.LostFocusCommand;
            string lostFocusCommandArgs = opts.LostFocusCommandArgs;
            string lostDocFocusCommand = opts.LostDocFocusCommand;
            string lostDocFocusCommandArgs = opts.LostDocFocusCommandArgs;
#endif

            saveRegistryKeyValue(SaveSingleDocumentSubKeyName, autosaveSingleDocEnabled.ToString());
            saveRegistryKeyValue(SaveFilesSubKeyName, autosaveEnabled.ToString());
            saveRegistryKeyValue(SaveProjectsSubKeyName, autosaveProjectEnabled.ToString());
            saveRegistryKeyValue(SaveSolutionsSubKeyName, autosaveSolutionEnabled.ToString());

#if USE_LOST_FOCUS_CMDS
            saveRegistryKeyValue(LostFocusCmdSubKeyName, lostFocusCommand);
            saveRegistryKeyValue(LostFocusCmdArgsSubKeyName, lostFocusCommandArgs);
            saveRegistryKeyValue(LostDocFocusCmdSubKeyName, lostDocFocusCommand);
            saveRegistryKeyValue(LostDocFocusCmdArgsSubKeyName, lostDocFocusCommandArgs);
#endif
            return base.QueryClose(out canClose);
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            //Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", ToString()));
            base.Initialize();

            try
            {
                _dte = (DTE2)GetService(typeof(DTE));
                vsShell = (IVsShell)GetService(typeof(SVsShell));
                Debug.Assert(vsShell != null, "vsShell is null");


                _textManager = (VsTextManager)GetService(typeof(VsTextManager));

                _componentModel = (IComponentModel)GetService(typeof(SComponentModel));
                _exportProvider = _componentModel.DefaultExportProvider;
                //_vim = _exportProvider.GetExportedValue<IVim>();
                //_vsAdapter = _exportProvider.GetExportedValue<IVsAdapter>();

                events = _dte.Events;
                Debug.Assert(events != null, "dte.Events is null");

                // Use the OutputWindowService to create the new pane.
                var oWinSvc = (IVsOutputWindow)GetService(typeof(SVsOutputWindow));
                Debug.Assert(oWinSvc != null, "oWinSvc is null");


                Int32 ret = oWinSvc.CreatePane(ref CBPkgPaneGuid, "CBExtensionPkg Output", 1, 0);
                Debug.Assert(ret == VSConstants.S_OK, "New pane not created");
                oWinSvc.GetPane(ref CBPkgPaneGuid, out mOutputWindowPane);
                Debug.Assert(mOutputWindowPane != null, "New pane not found");

                vsShell.AdviseBroadcastMessages(this, out abmCookie);

                // Retrieve the window events objects from the automation model.
                mWinEvents = events.WindowEvents[null];
                Debug.Assert(mWinEvents != null, "mWinEvents is null");
                mWinEvents.WindowActivated += mWinEvents_WindowActivated;

                if (InitFromRegistry() == false)
                {
                    // How to bail out successfully?
                }

                OutputMessage("CBExtensionPkg initialized");
            }
            catch (Exception e)
            {
                Debug.Assert(null != mOutputWindowPane);
                String s = string.Format("Initialize exception: {0} \n\t {1}", e.Source, e.Message);
                OutputMessage(s);
            }
        }

        private bool InitFromRegistry()
        {
            bool success = true;
            try
            {
                RegistryKey userKey = UserRegistryRoot;
                RegistryKey cbAddInKey = userKey.OpenSubKey(CBVSAddInSubKeyName, true);
                if (cbAddInKey == null)
                {
                    // todo: create it, populate it
                    cbAddInKey = userKey.CreateSubKey(CBVSAddInSubKeyName);
                    Debug.Assert(cbAddInKey != null, "CreateSubKey failed");
                    cbAddInKey.SetValue(SaveFilesSubKeyName, true);
                    cbAddInKey.SetValue(SaveProjectsSubKeyName, false);
                    cbAddInKey.Close();
                }

                cbAddInKey = userKey.OpenSubKey(CBVSAddInSubKeyName, true);
                Debug.Assert(cbAddInKey != null, "OpenSubKey failed");
                var saveFiles = cbAddInKey.GetValue(SaveFilesSubKeyName);
                var saveProjs = cbAddInKey.GetValue(SaveProjectsSubKeyName);
                var saveSolutions = cbAddInKey.GetValue(SaveSolutionsSubKeyName);
                var saveSingleDoc = cbAddInKey.GetValue(SaveSingleDocumentSubKeyName);
                var lostFocusCmd = cbAddInKey.GetValue(LostFocusCmdSubKeyName);
                var lostFocusCmdArgs = cbAddInKey.GetValue(LostFocusCmdArgsSubKeyName);


                if (!bool.TryParse((string)saveFiles, out bool autosaveEnabled))
                {
                    autosaveEnabled = true;
                }
                if (!bool.TryParse((string)saveProjs, out bool autosaveProjectEnabled))
                {
                    autosaveProjectEnabled = false;
                }
                if (!bool.TryParse((string)saveSolutions, out bool autosaveSolutionEnabled))
                {
                    autosaveSolutionEnabled = false;
                }
                if (!bool.TryParse((string)saveSingleDoc, out bool autosaveSingleDocEnabled))
                {
                    autosaveSingleDocEnabled = false;
                }

                var lostFocusCmdString = lostFocusCmd as string;
                if (string.IsNullOrWhiteSpace(lostFocusCmdString))
                {
                    lostFocusCmdString = "";
                }

                var lostFocusCmdArgsString = lostFocusCmdArgs as string;
                if (string.IsNullOrWhiteSpace(lostFocusCmdArgsString))
                {
                    lostFocusCmdArgsString = "";
                }

                var opts = GetDialogPage(typeof(AutoSaveOptions)) as AutoSaveOptions;
                Debug.Assert(null != opts, "OptionsPage obj is null");
                opts.AutoSaveDocuments = autosaveEnabled;
                opts.AutoSaveSingleDocument = autosaveSingleDocEnabled;
                opts.AutoSaveProjects = autosaveProjectEnabled;
                opts.AutoSaveSolution = autosaveSolutionEnabled;
#if USE_LOST_FOCUS_CMDS
                opts.LostFocusCommand = lostFocusCmdString;
                opts.LostFocusCommandArgs = lostFocusCmdArgsString;
#endif

                cbAddInKey.Close();
                userKey.Close();
                success = true;
            }
            catch (Exception)
            {
                success = false;
            }
            return success;
        }

        private void mWinEvents_WindowActivated(Window GotFocus, Window LostFocus)
        {
            string gotName = GotFocus != null ? GotFocus.Caption : "no GotFocus win";
            string lostName = LostFocus != null ? LostFocus.Caption : "no LostFocus win";
            DbgMessage($"GotFocus : {gotName}  LostFocus : {lostName}\n");
            if (LostFocus != null && LostFocus.Kind.Equals("Document"))
            {
                if (LostFocus.Document != null && !LostFocus.Document.Saved)
                {
                    var opts = GetDialogPage(typeof(AutoSaveOptions)) as AutoSaveOptions;
                    Debug.Assert(null != opts, "OptionsPage obj is null");

                    if (opts.AutoSaveSingleDocument)
                    {
                        //var currentDoc = _dte.ActiveDocument;
                        OutputMessage("   Saving...");
                        var sMsg = "    - " + LostFocus.Document.Name + "... " + "  ";
                        OutputMessage(sMsg);
                        vsSaveStatus stat = LostFocus.Document.Save();
                        OutputMessage(stat.ToString() + "\n");

                        // Can't be called in event handler
                        //LostFocus.Document.Activate();
                        //_dte.ExecuteCommand("File.SaveSelectedItems");
                        //currentDoc?.Activate();
                    }
                    else
                    {
                        OutputMessage("   There is a changed document, but Autosave Single Document not enabled...\n");
                    }
                }
            }

#if USE_LOST_FOCUS_CMDS
            if (GotFocus != null && GotFocus.Kind.Equals("Document"))
            {
                if (!_dte.ActiveDocument.Language.Equals("XAML"))
                {
                    if (!_dte.ActiveWindow.Caption.Contains(@"[Design]"))
                    {
                        ExecuteLostDocFocusCmd();
                    }
                }
            }
#endif
        }

#endregion

        private bool SafeExecuteCommand(ITextView contextTextView, string command, string args = "")
        {
            try
            {
                // Many Visual Studio commands expect focus to be in the editor when
                // running.  Switch focus there if an appropriate ITextView is available
                // var wpfTextView = contextTextView as IWpfTextView;
                var wpfTextView = contextTextView as IWpfTextView;
                if (wpfTextView != null)
                {
                    //wpfTextView.VisualElement.Focus();
                    //wpfTextView.VisualElement.Focus();
                }

                _dte.ExecuteCommand(command, args);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public int OnBroadcastMessage(uint msg, IntPtr wParam, IntPtr lParam)
        {
            var opts = GetDialogPage(typeof(AutoSaveOptions)) as AutoSaveOptions;
            Debug.Assert(null != opts, "OptionsPage obj is null");
            bool autosaveDocEnabled = opts.AutoSaveDocuments;
            bool autosaveProjectEnabled = opts.AutoSaveProjects;
            bool autosaveSolutionEnabled = opts.AutoSaveSolution;

            // DbgMessage("OnBroadcastMessage : msg: " + msg + " wParam : " + wParam + " lParam : " + lParam + "\n");
            // DbgMessage("   WM_ACTIVATEAPP == 0x001C\n");

            if (msg != WmActivateapp || 0 != (int)wParam) return 1;


            //
            // Save everything is enabled
            //
            if (autosaveDocEnabled && autosaveProjectEnabled && autosaveSolutionEnabled)
            {
                try
                {
                    _dte.ExecuteCommand("File.SaveAll");
#if USE_LOST_FOCUS_CMDS
                    ExecuteLostFocusCmd();
#endif
                    return 1;
                }
                catch (Exception)
                {
                    return 0;
                }
            }

            // List<Project> projects = new List<Project>();
            bool modifieldDocsPresent = false;
            bool needProjSave = false;

            // Test if any docs need saving
            foreach (Document document in _dte.Documents)
            {
                if (!document.Saved)
                {
                    DbgMessage(string.Format("  unsaved: {0}\n", document.Name));
                    modifieldDocsPresent = true;
                    break;
                }
            }

            if (modifieldDocsPresent)
            {
                if (autosaveDocEnabled)
                {
                    var currentDoc = _dte.ActiveDocument;

                    OutputMessage("   Saving...");
                    foreach (Document document in _dte.Documents)
                    {
                        bool docSaved = document.Saved;
                        if (!docSaved)
                        {
                            var sMsg = "    - " + document.Name + "... " + "  ";
                            OutputMessage(sMsg);
                            document.Activate();
                            _dte.ExecuteCommand("File.SaveSelectedItems");
                            // vsSaveStatus stat = document.Save();
                            //OutputMessage(stat.ToString() + "\n");
                        }
                    }
                    if (currentDoc != null)
                    {
                        currentDoc.Activate();
                    }
                }
                else
                {
                    OutputMessage("   There are changed docs, but Autosave documents not enabled...\n");
                }
            }
            else
            {
                OutputMessage("   No modified files to save...\n");
            }


            // Test if any projects need saving
            Projects Projects = GetCurrentProjects();
            foreach (Project project in Projects)
            {
                if (project.IsDirty)
                {
                    needProjSave = true;
                    break;
                }
            }

            if (needProjSave)
            {
                if (autosaveProjectEnabled)
                {
                    foreach (Project project in Projects)
                    {
                        if (project.IsDirty)
                        {
                            OutputMessage($"   Saving project file {project.Name}...");
                            project.Save();
                        }
                    }
                }
                else
                {
                    OutputMessage("   Project(s) changes, but Autosave projects not enabled...\n");
                }
            }
            else
            {
                OutputMessage("   No modified projects to save...\n");
            }

            if (_dte.Solution.IsDirty)
            {
                if (autosaveSolutionEnabled)
                {
                    OutputMessage($"   Saving solution file {_dte.Solution.FullName}...");
                    _dte.Solution.SaveAs(_dte.Solution.FileName);
                }
                else
                {
                    OutputMessage("   Solution changed, but Autosave solution not enabled...\n");
                }
            }

#if USE_LOST_FOCUS_CMDS
            ExecuteLostFocusCmd();
#endif

            return 1;
        }


#if USE_LOST_FOCUS_CMDS
        private void ExecuteLostFocusCmd()
        {
            var opts = GetDialogPage(typeof(AutoSaveOptions)) as AutoSaveOptions;
            Debug.Assert(null != opts, "OptionsPage obj is null");

            if (string.IsNullOrEmpty(opts?.LostFocusCommand)) { return; }

            if (_dte.Documents.Count > 0)
            {
                // this shouldn't be nec. -- too VsVim-specific.
                if (_dte.ActiveDocument != null)
                {
                    Debug.WriteLine($"     ActiveWindow: {_dte.ActiveWindow.Type}  {_dte.ActiveWindow.Kind}");
                    foreach (var win in _dte.Windows)
                    {
                        Debug.WriteLine($"   win: {win.GetType()}");
                    }
                    var typ = _dte.ActiveDocument.Type;
                    Debug.WriteLine($"    ActiveDoc type: {typ}");

                    try
                    {
                        //var thing = _textManager.GetActiveView()
                        // Execute command here, passing args.
                        DbgMessage("    Attempting to set to Normal mode.");
                        _dte.ExecuteCommand(opts.LostFocusCommand, opts.LostFocusCommandArgs);
                    }
                    catch (Exception xx)
                    {
                        OutputMessage($"Exception in trying to execute {opts.LostFocusCommand} {opts.LostFocusCommandArgs}: {xx.Message}");
                    }
                }
            }
        }

        private void ExecuteLostDocFocusCmd()
        {
            var opts = GetDialogPage(typeof(AutoSaveOptions)) as AutoSaveOptions;
            if (opts != null && !string.IsNullOrEmpty(opts.LostDocFocusCommand))
            {
                try
                {
                    // Execute command here, passing args.
                    DbgMessage("    Attempting to set to Normal mode.");
                    _dte.ExecuteCommand(opts.LostDocFocusCommand, opts.LostDocFocusCommandArgs);
                }
                catch (Exception xx)
                {
                    OutputMessage($"Exception in trying to execute {opts.LostDocFocusCommand} {opts.LostDocFocusCommandArgs}: {xx.Message}");
                }
            }
        }
#endif

        public Project GetCurrentProject()
        {
            Project dteProject = null;

            try
            {
                Array activeSolutionProjects = (Array)_dte.ActiveSolutionProjects;
                if (activeSolutionProjects == null)
                    throw new Exception("DTE.ActiveSolutionProjects returned null");

                if (activeSolutionProjects.Length > 0)
                {
                    dteProject = (Project)activeSolutionProjects.GetValue(0);
                    if (dteProject == null)
                        throw new Exception("DTE.ActiveSolutionProjects[0] returned null");
                }
            }
            catch (Exception xx)
            {
                Debug.WriteLine(xx.Message);
            }
            return dteProject;
        }


        public Projects GetCurrentProjects()
        {
            return _dte.Solution.Projects;
        }


        private void saveRegistryKeyValue(string keyName, string val)
        {
            RegistryKey userKey = UserRegistryRoot;
            RegistryKey cbAddInKey = userKey.OpenSubKey(CBVSAddInSubKeyName, true);
            Debug.Assert(cbAddInKey != null);
            cbAddInKey.SetValue(keyName, val);
            cbAddInKey.Close();
            userKey.Close();
        }

        private void ShowMsgBox(string s)
        {
            // Show a Message Box to prove we were here
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            if (uiShell == null)
            {
                return;
            }

            Guid clsid = Guid.Empty;
            ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                0,
                ref clsid,
                "CBExtensionPkg",
                string.Format(CultureInfo.CurrentCulture, s, ToString()),
                string.Empty,
                0,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                OLEMSGICON.OLEMSGICON_INFO,
                0, // false
                out int result));
        }

        private void OutputMessage(string msg)
        {
            // Make sure this service is available; might not be when VS is shutting down.
            var oWinSvc = (IVsOutputWindow)GetService(typeof(SVsOutputWindow));
            if (oWinSvc == null)
            {
                return;
            }

            if (mOutputWindowPane != null)
            {
                mOutputWindowPane.OutputString(msg);
                mOutputWindowPane.OutputString(Environment.NewLine);
            }
            else
            {
                ShowMsgBox(msg);
            }
        }

        private void DbgMessage(string msg)
        {
#if DEBUG
            OutputMessage(msg);
#endif
        }

        int IOleCommandTarget.Exec(ref Guid commandGroup, uint commandId, uint commandExecOpt, IntPtr variantIn, IntPtr variantOut)
        {
            Debug.WriteLine($"   IOleCommandTarget.Exec cmdGroup: {commandGroup} cmdId: {commandId}");
            return 0;
        }
        //private void ShowOptionsPage()
        //{
        //    IVsPackage vsPackage;
        //    _vsShell = serviceProvider.GetService<SVsShell, IVsShell>();
        //    Guid packageGuid = Constants.PackageGuid;
        //    if (ErrorHandler.Succeeded(_vsShell.LoadPackage(ref packageGuid, out vsPackage)))
        //    {
        //        var package = vsPackage as Package;
        //        if (package != null)
        //        {
        //            package.ShowOptionPage(typeof(AutoSaveOptions));
        //        }
        //    }
        //}
    }
}
