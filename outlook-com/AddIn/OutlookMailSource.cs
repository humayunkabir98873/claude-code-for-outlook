using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClaudeMailAssistant
{
    // Reads and acts on mail through Outlook's object model for one Outlook window: the main window (the
    // selected email, or a reply being written in the reading pane) or a window showing one email (read or
    // being written). It uses the Application object Outlook passed to the add-in, which Outlook trusts, so
    // reading sender addresses doesn't trigger the "a program is trying to access" prompt.
    sealed class OutlookMailSource : IMailSource, IDisposable
    {
        const string TransportHeaders = "http://schemas.microsoft.com/mapi/proptag/0x007D001F";
        const string SenderSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";

        readonly Outlook.Application app;
        readonly Outlook.Explorer explorer;   // set for the main window...
        readonly Outlook.Inspector inspector; // ...or for a window showing one item
        object inlineResponse;                // main window: the reply being written in the reading pane

        public event EventHandler SelectionChanged;

        public OutlookMailSource(Outlook.Application app, Outlook.Explorer explorer)
        {
            this.app = app;
            this.explorer = explorer;
            var events = (Outlook.ExplorerEvents_10_Event)explorer;
            events.SelectionChange += OnSelectionChange;
            events.InlineResponse += InlineResponseStarted;
            events.InlineResponseClose += InlineResponseClosed;
            inlineResponse = Try(() => explorer.ActiveInlineResponse);
        }

        public OutlookMailSource(Outlook.Application app, Outlook.Inspector inspector)
        {
            this.app = app;
            this.inspector = inspector;
        }

        public void Dispose()
        {
            if (explorer == null) return;
            var events = (Outlook.ExplorerEvents_10_Event)explorer;
            events.SelectionChange -= OnSelectionChange;
            events.InlineResponse -= InlineResponseStarted;
            events.InlineResponseClose -= InlineResponseClosed;
        }

        // Also called by Connect when it opens the pane because a reply was started in the reading pane.
        public void InlineResponseStarted(object item)
        {
            inlineResponse = item;
            OnSelectionChange();
        }

        void InlineResponseClosed()
        {
            inlineResponse = null;
            OnSelectionChange();
        }

        void OnSelectionChange()
        {
            try
            {
                SelectionChanged?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                Log.Error("SelectionChange", ex);
            }
        }

        public MailSnapshot Current
        {
            get
            {
                var draft = Draft();
                if (draft != null) return new MailSnapshot { Mode = "compose", EntryId = "", Subject = Try(() => (string)((dynamic)draft).Subject) ?? "" };

                var selected = inspector != null ? AsMessage(Try(() => inspector.CurrentItem)) : SelectedMessage();
                if (selected == null) return null;
                dynamic item = selected;
                return new MailSnapshot
                {
                    Mode = "read",
                    EntryId = (string)item.EntryID,
                    Subject = Try(() => (string)item.Subject) ?? "",
                    FromName = Try(() => (string)item.SenderName) ?? "",
                    FromAddress = SenderAddress(selected),
                    Created = Try(() => ((DateTime)item.ReceivedTime).ToString("s", CultureInfo.InvariantCulture)) ?? "",
                    Attachments = AttachmentNames(selected),
                };
            }
        }

        // The email being written in this window, if any.
        object Draft()
        {
            if (inspector == null) return inlineResponse;
            object item = Try(() => inspector.CurrentItem);
            return item is Outlook.MailItem mail && !mail.Sent ? item : null;
        }

        object SelectedMessage()
        {
            try
            {
                var selection = explorer.Selection;
                return selection.Count == 1 ? AsMessage(selection[1]) : null;
            }
            catch
            {
                return null; // some views, such as Outlook Today, have no selection
            }
        }

        static object AsMessage(object item) => item is Outlook.MailItem || item is Outlook.MeetingItem ? item : null;

        // An empty entryId means the draft in this window, which has no EntryID until it's saved.
        dynamic Item(string entryId) =>
            string.IsNullOrEmpty(entryId)
                ? Draft() ?? throw new InvalidOperationException("There's no email being written here.")
                : app.Session.GetItemFromID(entryId);

        public string GetBody(string entryId, bool html)
        {
            dynamic item = Item(entryId);
            string text = Try(() => (string)item.Body) ?? "";
            if (!html) return text;
            return Try(() => (string)item.HTMLBody) ?? "<pre>" + WebUtility.HtmlEncode(text) + "</pre>";
        }

        // Mail that never crossed the internet (such as internal Exchange mail) has no transport headers.
        public string GetHeaders(string entryId)
        {
            dynamic item = Item(entryId);
            return Try(() => (string)item.PropertyAccessor.GetProperty(TransportHeaders)) ?? "";
        }

        public void Reply(string entryId, string html)
        {
            dynamic reply = Item(entryId).Reply();
            // Display first so Outlook adds the user's signature, then put the draft above it.
            reply.Display();
            string existing = (string)reply.HTMLBody ?? "";
            string draft = "<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt\">" + html + "</div><br>";
            var bodyTag = Regex.Match(existing, "<body[^>]*>", RegexOptions.IgnoreCase);
            reply.HTMLBody = bodyTag.Success ? existing.Insert(bodyTag.Index + bodyTag.Length, draft) : draft + existing;
        }

        // Unlike the web add-in, the COM API can create the category, so it doesn't have to exist beforehand.
        public void AddCategory(string entryId, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            dynamic master = app.Session.Categories;
            bool known = false;
            for (int i = 1; i <= (int)master.Count && !known; i++)
                known = string.Equals((string)master.Item(i).Name, name, StringComparison.OrdinalIgnoreCase);
            if (!known) master.Add(name);

            dynamic item = Item(entryId);
            var current = ((string)item.Categories ?? "")
                .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(c => c.Trim())
                .Where(c => c.Length > 0)
                .ToList();
            if (current.Contains(name, StringComparer.OrdinalIgnoreCase)) return;
            current.Add(name);
            item.Categories = string.Join(CultureInfo.CurrentCulture.TextInfo.ListSeparator + " ", current);
            item.Save();
        }

        public void NewAppointment(string subject, DateTime start, DateTime end, string location, string body)
        {
            dynamic appointment = app.CreateItem(Outlook.OlItemType.olAppointmentItem);
            appointment.Subject = subject;
            appointment.Start = start;
            appointment.End = end;
            appointment.Location = location;
            appointment.Body = body;
            appointment.Display();
        }

        public void OpenItem(string entryId) => Item(entryId).Display();

        // ----- the draft being written -----

        public string GetSubject() => (string)Item("").Subject ?? "";

        public void SetSubject(string subject) => Item("").Subject = subject;

        public void SetBody(string html) => Item("").HTMLBody = html;

        // Outlook edits mail with Word, so the user's selection lives in the Word document.
        dynamic Selection()
        {
            object document = inspector != null ? inspector.WordEditor : explorer.ActiveInlineResponseWordEditor;
            if (document == null) throw new InvalidOperationException("The email editor isn't ready. Click in the email, then try again.");
            return ((dynamic)document).Windows.Item(1).Selection;
        }

        public string GetSelectedText()
        {
            dynamic selection = Selection();
            if ((int)selection.End <= (int)selection.Start) return "";
            // Word ends paragraphs with \r and manual line breaks with \v.
            return ((string)selection.Text ?? "").Replace('\r', '\n').Replace('\v', '\n');
        }

        public void ReplaceSelection(string text)
        {
            Selection().Text = (text ?? "").Replace("\r\n", "\n").Replace('\n', '\r');
        }

        // ----- helpers -----

        static string SenderAddress(object selected)
        {
            dynamic item = selected;
            if (Try(() => (string)item.SenderEmailType) == "EX")
            {
                return Try(() => (string)item.PropertyAccessor.GetProperty(SenderSmtpAddress))
                    ?? Try(() => (string)item.Sender.GetExchangeUser().PrimarySmtpAddress)
                    ?? "";
            }
            return Try(() => (string)item.SenderEmailAddress) ?? "";
        }

        static string[] AttachmentNames(object selected)
        {
            dynamic item = selected;
            var names = new List<string>();
            try
            {
                dynamic attachments = item.Attachments;
                for (int i = 1; i <= (int)attachments.Count; i++)
                {
                    dynamic attachment = attachments.Item(i);
                    names.Add(Try(() => (string)attachment.FileName) ?? Try(() => (string)attachment.DisplayName) ?? "attachment");
                }
            }
            catch (Exception ex)
            {
                Log.Error("AttachmentNames", ex);
            }
            return names.ToArray();
        }

        static T Try<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch
            {
                return default(T);
            }
        }
    }
}
