using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClaudeMailAssistant
{
    // Outlook loads this class at startup (Setup registers it under HKCU\Software\Microsoft\Office\Outlook\Addins).
    // It adds the Claude button and menu to the main window, opened emails, the compose window and the
    // right-click menu, gives each Outlook window its own Claude side pane, and opens that pane when the user
    // starts writing an email. Every callback catches its own exceptions: one escaping into Outlook gets the
    // add-in disabled.
    [ComVisible(true)]
    [Guid(Ids.ConnectClsid)]
    [ProgId(Ids.ConnectProgId)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        const string SettingsKey = @"HKEY_CURRENT_USER\Software\ClaudeMailAssistant";
        static readonly string InstallDir = Path.GetDirectoryName(typeof(Connect).Assembly.Location);

        static readonly Dictionary<string, string> RibbonFiles = new Dictionary<string, string>
        {
            ["Microsoft.Outlook.Explorer"] = "Ribbon.Explorer.xml",
            ["Microsoft.Outlook.Mail.Read"] = "Ribbon.Read.xml",
            ["Microsoft.Outlook.Mail.Compose"] = "Ribbon.Compose.xml",
        };

        sealed class Pane
        {
            public object Window; // Outlook.Explorer or Outlook.Inspector
            public Office.CustomTaskPane TaskPane;
            public TaskPane Control;
            public OutlookMailSource Mail;
        }

        readonly List<Pane> panes = new List<Pane>();
        readonly List<Office.IRibbonUI> ribbons = new List<Office.IRibbonUI>();
        readonly List<Outlook.Explorer> watchedExplorers = new List<Outlook.Explorer>(); // keeps their events alive
        Outlook.Application app;
        Outlook.Inspectors inspectors;
        Outlook.Explorers explorers;
        Office.ICTPFactory factory;

        // "Open Claude when I write an email", on unless the user turned it off in the Claude menu.
        static bool OpenWhenWriting
        {
            get => !(Registry.GetValue(SettingsKey, "OpenWhenWriting", null) is int value) || value != 0;
            set => Registry.SetValue(SettingsKey, "OpenWhenWriting", value ? 1 : 0, RegistryValueKind.DWord);
        }

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
                // Turned on from File > Options > Add-ins while Outlook runs: OnStartupComplete won't come.
                if (connectMode == ext_ConnectMode.ext_cm_AfterStartup) WatchWindows();
            }
            catch (Exception ex)
            {
                Log.Error("OnConnection", ex);
            }
        }

        public void OnStartupComplete(ref Array custom) => WatchWindows();

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            foreach (var pane in panes.ToArray()) Remove(pane);
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        // ----- opening the pane when the user starts writing -----

        void WatchWindows()
        {
            try
            {
                inspectors = app.Inspectors;
                ((Outlook.InspectorsEvents_Event)inspectors).NewInspector += OnNewInspector;
                explorers = app.Explorers;
                ((Outlook.ExplorersEvents_Event)explorers).NewExplorer += WatchExplorer;
                for (int i = 1; i <= explorers.Count; i++) WatchExplorer(explorers[i]);
            }
            catch (Exception ex)
            {
                Log.Error("WatchWindows", ex);
            }
        }

        void WatchExplorer(Outlook.Explorer explorer)
        {
            if (watchedExplorers.Contains(explorer)) return;
            watchedExplorers.Add(explorer);
            var events = (Outlook.ExplorerEvents_10_Event)explorer;
            events.InlineResponse += item => OnInlineResponse(explorer, item);
            events.Close += () => watchedExplorers.Remove(explorer);
        }

        // A new window: new mail, reply, forward or an opened email. Only drafts open the pane by themselves.
        void OnNewInspector(Outlook.Inspector inspector)
        {
            try
            {
                if (OpenWhenWriting && inspector.CurrentItem is Outlook.MailItem mail && !mail.Sent) Show(inspector);
            }
            catch (Exception ex)
            {
                Log.Error("NewInspector", ex);
            }
        }

        // Reply or forward written in the reading pane of the main window.
        void OnInlineResponse(Outlook.Explorer explorer, object item)
        {
            try
            {
                if (!OpenWhenWriting) return;
                bool existed = Find(explorer) != null;
                var pane = Show(explorer);
                if (!existed) pane.Mail.InlineResponseStarted(item); // the new pane missed this event
            }
            catch (Exception ex)
            {
                Log.Error("InlineResponse", ex);
            }
        }

        // ----- ribbon -----

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                return RibbonFiles.TryGetValue(ribbonId, out var file) ? Encoding.UTF8.GetString(WebHost.Resource(file)).TrimStart('﻿') : null;
            }
            catch (Exception ex)
            {
                Log.Error("GetCustomUI", ex);
                return null;
            }
        }

        public void CTPFactoryAvailable(Office.ICTPFactory ctpFactory) => factory = ctpFactory;

        // Callbacks below are called by name from the Ribbon.*.xml files.

        public void OnRibbonLoad(Office.IRibbonUI ribbon) => ribbons.Add(ribbon);

        public bool GetPressed(Office.IRibbonControl control)
        {
            try
            {
                return Find(WindowOf(control))?.TaskPane.Visible == true;
            }
            catch (Exception ex)
            {
                Log.Error("GetPressed", ex);
                return false;
            }
        }

        public void OnToggle(Office.IRibbonControl control, bool pressed)
        {
            Guard("OnToggle", () =>
            {
                var window = WindowOf(control);
                if (pressed) Show(window);
                else if (Find(window) is Pane pane) pane.TaskPane.Visible = false;
            });
        }

        // A menu item: open the pane and run the command named by its tag (see COMMANDS in main.js).
        public void OnCommand(Office.IRibbonControl control)
        {
            Guard("OnCommand", () => Show(WindowOf(control)).Control.RunAction(control.Tag));
        }

        public bool GetOpenWhenWriting(Office.IRibbonControl control) => OpenWhenWriting;

        public void OnOpenWhenWriting(Office.IRibbonControl control, bool pressed)
        {
            Guard("OnOpenWhenWriting", () =>
            {
                OpenWhenWriting = pressed;
                Invalidate();
            });
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

        // The ribbon passes its window; the right-click menu passes the selection, which belongs to the main window.
        object WindowOf(Office.IRibbonControl control)
        {
            object context = control.Context;
            return context is Outlook.Explorer || context is Outlook.Inspector ? context : app.ActiveExplorer();
        }

        void Invalidate()
        {
            foreach (var ribbon in ribbons)
            {
                try
                {
                    ribbon.Invalidate();
                }
                catch (Exception ex)
                {
                    Log.Error("Invalidate", ex);
                }
            }
        }

        static void Guard(string where, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Log.Error(where, ex);
                MessageBox.Show("Claude couldn't do that: " + ex.Message, Ids.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ----- panes -----

        Pane Show(object window)
        {
            var pane = Find(window) ?? Create(window);
            pane.TaskPane.Visible = true;
            Invalidate();
            return pane;
        }

        Pane Create(object window)
        {
            if (factory == null) throw new InvalidOperationException("Outlook hasn't made side panes available yet. Try again in a moment.");
            TaskPane.LastCreated = null;
            var ctp = factory.CreateCTP(Ids.PaneProgId, "Claude", window);
            var control = ctp.ContentControl as TaskPane ?? TaskPane.LastCreated
                ?? throw new InvalidOperationException("Outlook couldn't create the Claude pane.");
            ctp.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            ctp.Width = 420;

            var pane = new Pane { Window = window, TaskPane = ctp, Control = control };
            if (window is Outlook.Explorer explorer)
            {
                pane.Mail = new OutlookMailSource(app, explorer);
                ((Outlook.ExplorerEvents_10_Event)explorer).Close += () => Remove(pane);
            }
            else
            {
                var inspector = (Outlook.Inspector)window;
                pane.Mail = new OutlookMailSource(app, inspector);
                ((Outlook.InspectorEvents_10_Event)inspector).Close += () => Remove(pane);
            }
            ctp.VisibleStateChange += _ => Invalidate(); // un-press the button when the pane's X is used
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
        Pane Find(object window) => window == null ? null : panes.Find(p => p.Window == window);

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
