using System;
using System.Threading;
using System.Windows.Forms;

namespace ClaudeMailAssistant
{
    // Test only, never shipped: hosts the Claude pane in a plain window with made-up mail so the WebView2 +
    // office-shim + UI path can be exercised without Outlook. Every host call is printed. Type a command + Enter:
    //   read | none | compose      switch to the sample email, no selection, or a reply being written
    //   select <text>              pretend <text> is selected in the draft ("select" alone clears it)
    //   action <name>              a ribbon/right-click menu command, e.g. "action summarize"
    // Set WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9333 to drive the page over CDP.
    static class Harness
    {
        [STAThread]
        static void Main()
        {
            // Never share the installed add-in's browser profile, which a running Outlook may hold open.
            if (Environment.GetEnvironmentVariable("CLAUDE_MAIL_DATA_DIR") == null)
                Environment.SetEnvironmentVariable("CLAUDE_MAIL_DATA_DIR", System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ClaudeMailHarness"));
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
                    var parts = line.Trim().Split(new[] { ' ' }, 2);
                    var arg = parts.Length > 1 ? parts[1] : "";
                    form.BeginInvoke((Action)(() =>
                    {
                        switch (parts[0])
                        {
                            case "read": mail.Show(FakeMail.Mode.Read); break;
                            case "none": mail.Show(FakeMail.Mode.None); break;
                            case "compose": mail.Show(FakeMail.Mode.Compose); break;
                            case "select": mail.Selected = arg; break;
                            case "action": pane.RunAction(arg); break;
                        }
                    }));
                }
            }) { IsBackground = true }.Start();
            Application.Run(form);
        }
    }

    sealed class FakeMail : IMailSource
    {
        public enum Mode { None, Read, Compose }

        static readonly MailSnapshot Sample = new MailSnapshot
        {
            Mode = "read",
            EntryId = "00000000FAKEENTRYID",
            Subject = "Action required: verify your account",
            FromName = "PayPal paypal.com",
            FromAddress = "security@paypa1-alerts.xyz",
            Created = "2026-10-20T09:15:00",
            Attachments = new[] { "statement.pdf.zip" },
        };

        // A reply as Outlook (Word) writes it: the user's text, the signature bookmark, then the quoted email.
        const string ReplyHtml =
            "<html><head><style>p.MsoNormal{margin:0}</style></head><body lang=EN-US><div class=WordSection1>" +
            "<p class=MsoNormal>hi sam, thursday dont work for me can we do\r\nfriday instead<o:p></o:p></p>" +
            "<p class=MsoNormal><o:p>&nbsp;</o:p></p>" +
            "<p class=MsoNormal><a name=\"_MailAutoSig\"><span>Kabir Hossain</span></a></p><p class=MsoNormal>BD Writer</p>" +
            "<div><div style='border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in'>" +
            "<p class=MsoNormal><b>From:</b> Sam Lee<br><b>Sent:</b> Monday, October 19, 2026<br><b>Subject:</b> Review</p></div></div>" +
            "<p class=MsoNormal>Can we do the design review on Thursday at 3?</p></div></body></html>";

        Mode mode = Mode.Read;
        string subject = "RE: Review";
        string body = ReplyHtml;
        public string Selected = "";

        public event EventHandler SelectionChanged;

        public void Show(Mode next)
        {
            mode = next;
            Trace("show " + next);
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }

        public MailSnapshot Current =>
            mode == Mode.Read ? Sample : mode == Mode.Compose ? new MailSnapshot { Mode = "compose", EntryId = "", Subject = subject } : null;

        public string GetBody(string entryId, bool html)
        {
            Trace($"GetBody '{entryId}' html={html}");
            if (mode == Mode.Compose) return html ? body : "(plain text of the draft)";
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

        public string GetSubject() => subject;

        public void SetSubject(string value)
        {
            Trace("SetSubject " + value);
            subject = value;
        }

        public void SetBody(string html)
        {
            Trace("SetBody " + html);
            body = html;
        }

        public string GetSelectedText() => Selected;

        public void ReplaceSelection(string text) => Trace("ReplaceSelection " + text);

        static void Trace(string message) => Console.WriteLine("[host] " + message);
    }
}
