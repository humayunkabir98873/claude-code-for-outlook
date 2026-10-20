using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace ClaudeMailAssistant
{
    // Test only, never shipped: hosts the Claude pane in a plain window with a made-up email so the
    // WebView2 + office-shim + UI path can be exercised without Outlook. Every host call is printed.
    //   Harness.exe            then type "none" or "sample" + Enter to change the selection
    // Set WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9333 to drive the page over CDP.
    static class Harness
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            var form = new Form { Text = "Claude pane harness", Width = 460, Height = 900 };
            var pane = new TaskPane { Dock = DockStyle.Fill };
            var mail = new FakeMail();
            form.Controls.Add(pane);
            form.Load += (s, e) => pane.Attach(mail);
            new Thread(() =>
            {
                string line;
                while ((line = Console.ReadLine()) != null)
                {
                    var command = line.Trim();
                    form.BeginInvoke((Action)(() => mail.Select(command == "none" ? null : FakeMail.Sample)));
                }
            }) { IsBackground = true }.Start();
            Application.Run(form);
        }
    }

    sealed class FakeMail : IMailSource
    {
        public static readonly MailSnapshot Sample = new MailSnapshot
        {
            EntryId = "00000000FAKEENTRYID",
            Subject = "Action required: verify your account",
            FromName = "PayPal paypal.com",
            FromAddress = "security@paypa1-alerts.xyz",
            Created = "2026-10-20T09:15:00",
            Attachments = new[] { "statement.pdf.zip" },
        };

        MailSnapshot current = Sample;

        public event EventHandler SelectionChanged;

        public MailSnapshot Current => current;

        public void Select(MailSnapshot snapshot)
        {
            current = snapshot;
            Trace("select " + (snapshot?.Subject ?? "(none)"));
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }

        public string GetBody(string entryId, bool html)
        {
            Trace($"GetBody {entryId} html={html}");
            return html
                ? "<p>Dear customer,</p><p>Your account is limited. <a href=\"http://198.51.100.7/login\">https://www.paypal.com/signin</a></p>"
                : "Dear customer,\r\nYour account is limited. Verify within 24 hours at https://www.paypal.com/signin";
        }

        public string GetHeaders(string entryId)
        {
            Trace("GetHeaders " + entryId);
            return "Authentication-Results: mx.example; spf=fail smtp.mailfrom=paypa1-alerts.xyz; dkim=none; dmarc=fail\r\nReply-To: help@other-domain.ru\r\n";
        }

        public void Reply(string entryId, string html) => Trace($"Reply {entryId}: {html}");

        public void AddCategory(string entryId, string name) => Trace($"AddCategory {entryId}: {name}");

        public void NewAppointment(string subject, DateTime start, DateTime end, string location, string body) =>
            Trace($"NewAppointment {subject} {start:s}..{end:s} at {location}");

        public void OpenItem(string entryId) => Trace("OpenItem " + entryId);

        static void Trace(string message) => Console.WriteLine("[host] " + message);
    }
}
