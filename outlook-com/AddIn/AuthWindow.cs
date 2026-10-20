using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace ClaudeMailAssistant
{
    // The window graph.js opens through Office.context.ui.displayDialogAsync: auth.html signs in to Microsoft
    // with MSAL and hands the token back to the pane through Office.context.ui.messageParent.
    sealed class AuthWindow : Form
    {
        readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
        Bridge bridge;

        public event Action<string> MessageReceived;
        public bool ClosedByAddIn;

        public AuthWindow(string url, int widthPercent, int heightPercent)
        {
            Text = "Sign in to Microsoft - " + Ids.ProductName;
            StartPosition = FormStartPosition.CenterScreen;
            var area = Screen.PrimaryScreen.WorkingArea;
            Size = new Size(Math.Max(520, area.Width * widthPercent / 100), Math.Max(640, area.Height * heightPercent / 100));
            Controls.Add(web);
            Load += async (s, e) =>
            {
                try
                {
                    await WebHost.InitializeAsync(web, ownPagesOnly: false); // Microsoft's sign-in pages load here
                    bridge = new Bridge(web.CoreWebView2, HandleRequest);
                    web.CoreWebView2.Navigate(url);
                }
                catch (Exception ex)
                {
                    Log.Error("AuthWindow", ex);
                    MessageReceived?.Invoke(Bridge.Serialize(new Dictionary<string, object> { ["error"] = "The sign-in window couldn't start: " + ex.Message }));
                    ClosedByAddIn = true;
                    Close();
                }
            };
        }

        object HandleRequest(string op, Dictionary<string, object> args)
        {
            switch (op)
            {
                case "hello":
                    return new Dictionary<string, object> { ["role"] = "dialog" };
                case "messageParent":
                    MessageReceived?.Invoke(Bridge.Text(args, "message"));
                    return null;
                default:
                    throw new BridgeException(0, "Not available in the sign-in window: " + op);
            }
        }
    }
}
