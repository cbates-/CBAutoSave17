// Project:  CBExtensionPkg
//


using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
// ReSharper disable LocalizableElement
// ReSharper disable ConvertToAutoProperty

namespace BlackIceSoftware.CBExtensionPkg
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    internal class AutoSaveOptions : DialogPage
    {
        private bool _autoSaveDocuments;
        private bool _autoSaveSingleDocument;

        private bool _autoSaveProjects;
        private bool _autoSaveSolution;

#if USE_LOST_FOCUS_CMDS
        private string _lostFocusCommand;
        private string _lostFocusCommandArgs;
        private string _lostDocFocusCommand;
        private string _lostDocFocusCommandArgs;
#endif

        [Category("General")]
        [DisplayName("AutoSaveDocuments")]
        [Description("Save changed documents when VS loses focus")]
        public bool AutoSaveDocuments
        {
            get { return _autoSaveDocuments; }
            set { _autoSaveDocuments = value; }
        }

        [Category("General")]
        [DisplayName("AutoSaveSingleDocument")]
        [Description("Save changed document when document window loses focus")]
        public bool AutoSaveSingleDocument
        {
            get { return _autoSaveSingleDocument; }
            set { _autoSaveSingleDocument = value; }
        }

        [Category("General")]
        [DisplayName("AutoSaveProjects")]
        [Description("Save changed projects when VS loses focus")]
        public bool AutoSaveProjects
        {
            get { return _autoSaveProjects; }
            set { _autoSaveProjects = value; }
        }

        [Category("General")]
        [DisplayName("AutoSaveSolution")]
        [Description("Save changed solution when VS loses focus")]
        public bool AutoSaveSolution
        {
            get { return _autoSaveSolution; }
            set { _autoSaveSolution = value; }
        }

#if USE_LOST_FOCUS_CMDS
        [Category("Custom")]
        [DisplayName("LostFocusCommand")]
        [Description("Command to execute when VisStudio loses focus - USE WITH CAUTION")]
        public string LostFocusCommand
        {
            get { return _lostFocusCommand; }
            set { _lostFocusCommand = value; }
        }

        [Category("Custom")]
        [DisplayName("LostFocusCommandArgs")]
        [Description("Command args for the LostFocusCommand")]
        public string LostFocusCommandArgs
        {
            get { return _lostFocusCommandArgs; }
            set { _lostFocusCommandArgs = value; }
        }

        [Category("Custom")]
        [DisplayName("LostDocFocusCommand")]
        [Description("Command to execute when a Document window loses focus - USE WITH CAUTION")]
        public string LostDocFocusCommand
        {
            get { return _lostDocFocusCommand; }
            set { _lostDocFocusCommand = value; }
        }

        [Category("Custom")]
        [DisplayName("LostDocFocusCommandArgs")]
        [Description("Command args for the LostDocFocusCommand")]
        public string LostDocFocusCommandArgs
        {
            get { return _lostDocFocusCommandArgs; }
            set { _lostDocFocusCommandArgs = value; }
        }
#endif

    }
}