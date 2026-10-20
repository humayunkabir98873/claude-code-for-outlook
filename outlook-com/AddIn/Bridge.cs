using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using Microsoft.Web.WebView2.Core;

namespace ClaudeMailAssistant
{
    // JSON request/response channel between a page's office-shim.js and C#:
    //   page -> host  {id, op, args}        host -> page  {id, result} or {id, error: {code, message}}
    //   host -> page  {event, ...}          (selection changes, sign-in window messages)
    // Only our own virtual-host pages are answered, so a site shown in the sign-in window can't drive Outlook.
    sealed class Bridge
    {
        static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        readonly CoreWebView2 core;
        readonly Func<string, Dictionary<string, object>, object> handle;

        public Bridge(CoreWebView2 core, Func<string, Dictionary<string, object>, object> handle)
        {
            this.core = core;
            this.handle = handle;
            core.WebMessageReceived += OnMessage;
        }

        public static string Serialize(object value) => Json.Serialize(value);

        public void Post(object message)
        {
            try
            {
                core.PostWebMessageAsJson(Json.Serialize(message));
            }
            catch (Exception ex)
            {
                Log.Error("Bridge.Post", ex);
            }
        }

        void OnMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            if (!WebHost.IsOwnPage(e.Source)) return;
            object id = null;
            try
            {
                var message = Json.Deserialize<Dictionary<string, object>>(e.WebMessageAsJson);
                message.TryGetValue("id", out id);
                var args = message.TryGetValue("args", out var a) ? a as Dictionary<string, object> : null;
                var result = handle(Text(message, "op"), args ?? new Dictionary<string, object>());
                Post(new Dictionary<string, object> { ["id"] = id, ["result"] = result });
            }
            catch (Exception ex)
            {
                Log.Error("Bridge", ex);
                var error = new Dictionary<string, object> { ["code"] = (ex as BridgeException)?.Code ?? 0, ["message"] = Describe(ex) };
                Post(new Dictionary<string, object> { ["id"] = id, ["error"] = error });
            }
        }

        static string Describe(Exception ex) =>
            ex is COMException ? "Outlook couldn't do that: " + ex.Message : ex.Message;

        public static string Text(Dictionary<string, object> args, string key) =>
            args.TryGetValue(key, out var v) && v != null ? Convert.ToString(v, CultureInfo.InvariantCulture) : "";

        public static int Number(Dictionary<string, object> args, string key, int fallback) =>
            args.TryGetValue(key, out var v) && v != null ? Convert.ToInt32(v, CultureInfo.InvariantCulture) : fallback;

        public static bool Flag(Dictionary<string, object> args, string key) =>
            args.TryGetValue(key, out var v) && v is bool b && b;

        public static IEnumerable<string> List(Dictionary<string, object> args, string key) =>
            args.TryGetValue(key, out var v) && v is IEnumerable items && !(v is string)
                ? items.Cast<object>().Select(o => Convert.ToString(o, CultureInfo.InvariantCulture)).ToArray()
                : new string[0];
    }

    // An error with an Office dialog error code (12004, 12007, ...) that graph.js already knows how to explain.
    sealed class BridgeException : Exception
    {
        public readonly int Code;

        public BridgeException(int code, string message) : base(message)
        {
            Code = code;
        }
    }
}
