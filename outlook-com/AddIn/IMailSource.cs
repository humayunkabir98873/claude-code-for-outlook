using System;

namespace ClaudeMailAssistant
{
    // What the task pane needs from Outlook. OutlookMailSource is the real one; the test harness fakes it.
    public interface IMailSource
    {
        event EventHandler SelectionChanged;

        // The selected email, or null when nothing (or not exactly one message) is selected.
        MailSnapshot Current { get; }

        string GetBody(string entryId, bool html);
        string GetHeaders(string entryId);
        void Reply(string entryId, string html);
        void AddCategory(string entryId, string name);
        void NewAppointment(string subject, DateTime start, DateTime end, string location, string body);
        void OpenItem(string entryId);
    }

    public sealed class MailSnapshot
    {
        public string EntryId;
        public string Subject;
        public string FromName;
        public string FromAddress;
        public string Created; // ISO 8601, local time
        public string[] Attachments;
    }
}
