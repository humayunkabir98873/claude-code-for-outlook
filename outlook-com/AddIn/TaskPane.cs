using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ClaudeMailAssistant
{
    // The Claude side pane: an ActiveX control created by Office's custom task pane factory. It shows
    // public/taskpane.html in WebView2 and answers office-shim.js requests from an IMailSource.
    [ComVisible(true)]
    [Guid(Ids.PaneClsid)]
    [ProgId(Ids.PaneProgId)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TaskPane : UserControl
    {
        // Office creates the control through COM; Connect falls back to this if ContentControl can't be cast.
        internal static TaskPane LastCreated;

        readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
        IMailSource mail;
        Bridge bridge;
        AuthWindow dialog;

        public TaskPane()
        {
            Controls.Add(web);
            LastCreated = this;
        }

        public async void Attach(IMailSource source)
        {
            mail = source;
            mail.SelectionChanged += (s, e) =>
                bridge?.Post(new Dictionary<string, object> { ["event"] = "itemChanged", ["item"] = ToJson(mail.Current) });
            try
            {
                await WebHost.InitializeAsync(web, ownPagesOnly: true);
                bridge = new Bridge(web.CoreWebView2, HandleRequest);
                web.CoreWebView2.Navigate(WebHost.Url("taskpane.html"));
            }
            catch (Exception ex)
            {
                Log.Error("TaskPane.Attach", ex);
                ShowFailure(ex);
            }
        }

        object HandleRequest(string op, Dictionary<string, object> args)
        {
            string entryId = Bridge.Text(args, "entryId");
            switch (op)
            {
                case "hello":
                    return new Dictionary<string, object> { ["role"] = "pane", ["item"] = ToJson(mail.Current) };
                case "getBody":
                    return mail.GetBody(entryId, Bridge.Flag(args, "html"));
                case "getHeaders":
                    return mail.GetHeaders(entryId);
                case "reply":
                    mail.Reply(entryId, Bridge.Text(args, "html"));
                    return null;
                case "addCategory":
                    foreach (var name in Bridge.List(args, "names")) mail.AddCategory(entryId, name);
                    return null;
                case "newAppointment":
                    mail.NewAppointment(Bridge.Text(args, "subject"), LocalTime(Bridge.Text(args, "start")), LocalTime(Bridge.Text(args, "end")),
                        Bridge.Text(args, "location"), Bridge.Text(args, "body"));
                    return null;
                case "openItem":
                    mail.OpenItem(entryId);
                    return null;
                case "openDialog":
                    OpenDialog(Bridge.Text(args, "url"), Bridge.Number(args, "width", 30), Bridge.Number(args, "height", 60));
                    return null;
                case "closeDialog":
                    CloseDialog();
                    return null;
                case "showError":
                    // Never block inside a WebView2 event handler; show the message once it returns.
                    var text = Bridge.Text(args, "message");
                    BeginInvoke((Action)(() => MessageBox.Show(this, text, Ids.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)));
                    return null;
                default:
                    throw new BridgeException(0, "Unknown request: " + op);
            }
        }

        void OpenDialog(string url, int widthPercent, int heightPercent)
        {
            if (!WebHost.IsOwnPage(url)) throw new BridgeException(12004, "The sign-in window can only open the add-in's own pages.");
            if (dialog != null) throw new BridgeException(12007, "A sign-in window is already open.");
            var window = new AuthWindow(url, widthPercent, heightPercent);
            window.MessageReceived += message =>
                bridge.Post(new Dictionary<string, object> { ["event"] = "dialogMessage", ["message"] = message });
            window.FormClosed += (s, e) =>
            {
                if (!window.ClosedByAddIn)
                    bridge.Post(new Dictionary<string, object> { ["event"] = "dialogEvent", ["error"] = 12006 });
                dialog = null;
            };
            dialog = window;
            window.Show(this);
        }

        void CloseDialog()
        {
            if (dialog == null) return;
            dialog.ClosedByAddIn = true;
            dialog.Close();
        }

        void ShowFailure(Exception ex)
        {
            Controls.Clear();
            Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                Text = ex is WebView2RuntimeNotFoundException
                    ? "Claude needs the Microsoft Edge WebView2 Runtime. Install it from https://go.microsoft.com/fwlink/p/?LinkId=2124703, then restart Outlook."
                    : "Claude couldn't start: " + ex.Message + "\r\n\r\nDetails: " + Path.Combine(Log.DataDir, "log.txt"),
            });
        }

        static DateTime LocalTime(string iso) => DateTimeOffset.Parse(iso, CultureInfo.InvariantCulture).LocalDateTime;

        static Dictionary<string, object> ToJson(MailSnapshot s) => s == null ? null : new Dictionary<string, object>
        {
            ["entryId"] = s.EntryId,
            ["subject"] = s.Subject,
            ["fromName"] = s.FromName,
            ["fromAddress"] = s.FromAddress,
            ["created"] = s.Created,
            ["attachments"] = s.Attachments ?? new string[0],
        };

        protected override void Dispose(bool disposing)
        {
            if (disposing) dialog?.Dispose();
            base.Dispose(disposing);
        }
    }
}
