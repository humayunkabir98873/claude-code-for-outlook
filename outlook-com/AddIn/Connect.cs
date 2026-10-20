using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClaudeMailAssistant
{
    // Outlook loads this class at startup (Setup registers it under HKCU\Software\Microsoft\Office\Outlook\Addins).
    // It adds the Claude button to the Home tab and gives each Outlook window its own Claude side pane.
    // Every callback catches its own exceptions: one escaping into Outlook gets the add-in disabled.
    [ComVisible(true)]
    [Guid(Ids.ConnectClsid)]
    [ProgId(Ids.ConnectProgId)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        const string ToggleId = "ClaudeToggle";
        static readonly string InstallDir = Path.GetDirectoryName(typeof(Connect).Assembly.Location);

        sealed class Pane
        {
            public Outlook.Explorer Explorer;
            public Office.CustomTaskPane TaskPane;
            public OutlookMailSource Mail;
        }

        readonly List<Pane> panes = new List<Pane>();
        Outlook.Application app;
        Office.ICTPFactory factory;
        Office.IRibbonUI ribbon;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                AppDomain.CurrentDomain.AssemblyResolve += ResolveFromInstallDir;
                // WebView2's async setup must resume on Outlook's UI thread.
                if (!(SynchronizationContext.Current is WindowsFormsSynchronizationContext))
                    SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                app = (Outlook.Application)application;
                Log.Write($"Loaded {Ids.Version} into Outlook {app.Version} ({(Environment.Is64BitProcess ? "64" : "32")}-bit)");
            }
            catch (Exception ex)
            {
                Log.Error("OnConnection", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            foreach (var pane in panes.ToArray()) Remove(pane);
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                return ribbonId == "Microsoft.Outlook.Explorer" ? Encoding.UTF8.GetString(WebHost.Resource("Ribbon.xml")).TrimStart('﻿') : null;
            }
            catch (Exception ex)
            {
                Log.Error("GetCustomUI", ex);
                return null;
            }
        }

        public void CTPFactoryAvailable(Office.ICTPFactory ctpFactory) => factory = ctpFactory;

        // ----- ribbon callbacks, called by name from Ribbon.xml -----

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI) => ribbon = ribbonUI;

        public bool GetPressed(Office.IRibbonControl control)
        {
            try
            {
                return Find(control.Context as Outlook.Explorer)?.TaskPane.Visible == true;
            }
            catch (Exception ex)
            {
                Log.Error("GetPressed", ex);
                return false;
            }
        }

        public void OnToggle(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                var explorer = control.Context as Outlook.Explorer;
                if (explorer == null) return;
                var pane = Find(explorer) ?? (pressed ? Create(explorer) : null);
                if (pane != null) pane.TaskPane.Visible = pressed;
            }
            catch (Exception ex)
            {
                Log.Error("OnToggle", ex);
                MessageBox.Show("Claude couldn't open its pane: " + ex.Message, Ids.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public object LoadImage(string imageId)
        {
            try
            {
                using (var stream = new MemoryStream(WebHost.Resource("icon-80.png")))
                using (var image = new Bitmap(stream))
                    return PictureConverter.ToPicture(image);
            }
            catch (Exception ex)
            {
                Log.Error("LoadImage", ex);
                return null;
            }
        }

        // ----- panes -----

        Pane Create(Outlook.Explorer explorer)
        {
            if (factory == null) throw new InvalidOperationException("Outlook hasn't made side panes available yet. Try again in a moment.");
            TaskPane.LastCreated = null;
            var ctp = factory.CreateCTP(Ids.PaneProgId, "Claude", explorer);
            var control = ctp.ContentControl as TaskPane ?? TaskPane.LastCreated
                ?? throw new InvalidOperationException("Outlook couldn't create the Claude pane.");
            ctp.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            ctp.Width = 420;

            var pane = new Pane { Explorer = explorer, TaskPane = ctp, Mail = new OutlookMailSource(app, explorer) };
            ctp.VisibleStateChange += _ =>
            {
                try
                {
                    ribbon?.InvalidateControl(ToggleId); // un-press the button when the pane's X is used
                }
                catch (Exception ex)
                {
                    Log.Error("VisibleStateChange", ex);
                }
            };
            ((Outlook.ExplorerEvents_10_Event)explorer).Close += () => Remove(pane);
            panes.Add(pane);
            control.Attach(pane.Mail);
            return pane;
        }

        void Remove(Pane pane)
        {
            try
            {
                panes.Remove(pane);
                pane.Mail.Dispose();
                pane.TaskPane.Delete();
            }
            catch (Exception ex)
            {
                Log.Error("Remove", ex);
            }
        }

        // Outlook hands back the same wrapper object for the same window, so reference equality finds it.
        Pane Find(Outlook.Explorer explorer) => explorer == null ? null : panes.Find(p => p.Explorer == explorer);

        // Outlook's own folder is the probing path, so load WebView2's assemblies from ours.
        static Assembly ResolveFromInstallDir(object sender, ResolveEventArgs e)
        {
            var file = Path.Combine(InstallDir, new AssemblyName(e.Name).Name + ".dll");
            return File.Exists(file) ? Assembly.LoadFrom(file) : null;
        }

        // AxHost exposes the Image-to-IPictureDisp conversion the ribbon needs for a custom icon.
        sealed class PictureConverter : AxHost
        {
            PictureConverter() : base(string.Empty) { }

            public static object ToPicture(Image image) => GetIPictureDispFromPicture(image);
        }
    }
}
