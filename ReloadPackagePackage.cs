using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using OLE = Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Extensibility;
using System.IO;
using Microsoft.VisualStudio.OLE.Interop;

namespace ScottHanselman.ReloadPackage
{
    internal class NativeMethods
    {
        [DllImport("Ole32.dll", EntryPoint = "CreateStreamOnHGlobal")]
        internal static extern void CreateStreamOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease, [Out] out Microsoft.VisualStudio.OLE.Interop.IStream stream);
    }

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
    // This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidReloadPackagePkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    //[ProvideAutoLoad("{ADFC4E64-0397-11D1-9F4E-00A0C911004F}")]
    //[ProvideAutoLoad("{F1536EF8-92EC-443C-9ED7-FDADF150DA82}")]
    //[ProvideAutoLoad("{ADFC4E62-0397-11D1-9F4E-00A0C911004F}")]
    //[ProvideAutoLoad("{ADFC4E63-0397-11D1-9F4E-00A0C911004F}")]
    //[ProvideAutoLoad("{ADFC4E61-0397-11D1-9F4E-00A0C911004F}")]
    public sealed class ReloadPackagePackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ReloadPackagePackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overriden Package Implementation
        #region Package Members

        SolutionEventsListener listener = null;
        IVsUIShellDocumentWindowMgr winmgr = null;
        IStream comStream;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initilaization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));

            //SolutionEventsListener class from http://stackoverflow.com/questions/2525457/automating-visual-studio-with-envdte 
            // via Elisha http://stackoverflow.com/users/167149/elisha
            listener = new SolutionEventsListener();

            winmgr = Package.GetGlobalService(typeof(IVsUIShellDocumentWindowMgr)) as IVsUIShellDocumentWindowMgr;

            listener.OnQueryUnloadProject += () =>
            {
                Debug.WriteLine("HANSELMAN: Before Unload Project!");
                comStream = SaveDocumentWindowPositions(winmgr);
            };
            listener.OnAfterOpenProject += () => { 
                int hr = winmgr.ReopenDocumentWindows(comStream);
                comStream = null;
                Debug.WriteLine(String.Format("HANSELMAN: After Project Loaded! hr=", hr));
            };

            base.Initialize();

        }

        #endregion

        private IStream SaveDocumentWindowPositions(IVsUIShellDocumentWindowMgr windowsMgr)
        {
            if (windowsMgr == null)
            {
                Debug.Assert(false, "IVsUIShellDocumentWindowMgr", String.Empty, 0);
                return null;
            }
            IStream stream;
            NativeMethods.CreateStreamOnHGlobal(IntPtr.Zero, true, out stream);
            if (stream == null)
            {
                Debug.Assert(false, "CreateStreamOnHGlobal", String.Empty, 0);
                return null;
            }
            int hr = windowsMgr.SaveDocumentWindowPositions(0, stream);
            if (hr != VSConstants.S_OK)
            {
                Debug.Assert(false, "SaveDocumentWindowPositions", String.Empty, hr);
                return null;
            }

            // Move to the beginning of the stream 
            // In preparation for reading
            LARGE_INTEGER l = new LARGE_INTEGER();
            ULARGE_INTEGER[] ul = new ULARGE_INTEGER[1];
            ul[0] = new ULARGE_INTEGER();
            l.QuadPart = 0;
            //Seek to the beginning of the stream
            stream.Seek(l, 0, ul);
            return stream;
        }

    }

    public class SolutionEventsListener : IVsSolutionEvents, IDisposable
    {
        private IVsSolution solution;
        private uint solutionEventsCookie;

        public event Action OnAfterOpenProject;
        public event Action OnQueryUnloadProject;


        public SolutionEventsListener()
        {
            InitNullEvents();

            solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;

            if (solution != null)
            {
                solution.AdviseSolutionEvents(this, out solutionEventsCookie);
            }
        }

        private void InitNullEvents()
        {
            OnAfterOpenProject += () => { };
            OnQueryUnloadProject += () => { };
        }

        #region IVsSolutionEvents Members

        int IVsSolutionEvents.OnAfterCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            OnAfterOpenProject();
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            OnQueryUnloadProject();
            return VSConstants.S_OK;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (solution != null && solutionEventsCookie != 0)
            {
                GC.SuppressFinalize(this);
                solution.UnadviseSolutionEvents(solutionEventsCookie);
                OnQueryUnloadProject = null;
                OnAfterOpenProject = null;
                solutionEventsCookie = 0;
                solution = null;
            }
        }

        #endregion
    }

}
