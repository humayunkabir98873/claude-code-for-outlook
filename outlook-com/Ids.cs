namespace ClaudeMailAssistant
{
    // Identities shared by the add-in and its installer. Changing a GUID or ProgID orphans existing installs.
    static class Ids
    {
        public const string ProductName = "Claude Mail Assistant";
        public const string Version = "1.0.0";

        public const string ConnectProgId = "ClaudeMailAssistant.Connect";
        public const string ConnectClsid = "02FBE1B4-C7B9-4452-BA5E-97B78DD5B8FA";
        public const string PaneProgId = "ClaudeMailAssistant.TaskPane";
        public const string PaneClsid = "B2BE7BC2-F80E-4256-A425-7FE78920140A";

        // public/ is served from this made-up host inside WebView2; .example never resolves on a real network.
        public const string VirtualHost = "claude-mail-assistant.example";
    }
}
