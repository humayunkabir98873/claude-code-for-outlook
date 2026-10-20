using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ClaudeMailAssistant
{
    // Shared WebView2 setup: one browser environment (so the pane and the sign-in window share settings and
    // storage), public/ served from a virtual host, and office.js swapped for office-shim.js.
    static class WebHost
    {
        public const string Origin = "https://" + Ids.VirtualHost;
        public static readonly string InstallDir = Path.GetDirectoryName(typeof(WebHost).Assembly.Location);

        static Task<CoreWebView2Environment> environment;
        static byte[] shim;

        public static string Url(string page) => Origin + "/" + page;

        public static bool IsOwnPage(string uri) =>
            uri != null && uri.StartsWith(Origin + "/", StringComparison.OrdinalIgnoreCase);

        public static byte[] Resource(string name)
        {
            using (var stream = typeof(WebHost).Assembly.GetManifestResourceStream(name))
            using (var copy = new MemoryStream())
            {
                if (stream == null) throw new FileNotFoundException("Missing embedded resource " + name);
                stream.CopyTo(copy);
                return copy.ToArray();
            }
        }

        static Task<CoreWebView2Environment> SharedEnvironment()
        {
            if (environment == null)
            {
                CoreWebView2Environment.SetLoaderDllFolderPath(Path.Combine(InstallDir, Environment.Is64BitProcess ? "x64" : "x86"));
                environment = CoreWebView2Environment.CreateAsync(null, Path.Combine(Log.DataDir, "WebView2"));
            }
            return environment;
        }

        // ownPagesOnly: links to other sites open in the default browser instead of inside the pane.
        public static async Task InitializeAsync(WebView2 web, bool ownPagesOnly)
        {
            CoreWebView2Environment env;
            try
            {
                env = await SharedEnvironment();
            }
            catch
            {
                environment = null; // let the next pane retry
                throw;
            }
            await web.EnsureCoreWebView2Async(env);

            var core = web.CoreWebView2;
            core.SetVirtualHostNameToFolderMapping(Ids.VirtualHost, Path.Combine(InstallDir, "public"), CoreWebView2HostResourceAccessKind.DenyCors);
            core.AddWebResourceRequestedFilter("https://appsforoffice.microsoft.com/*", CoreWebView2WebResourceContext.All);
            core.WebResourceRequested += (s, e) =>
            {
                shim = shim ?? Resource("office-shim.js");
                e.Response = env.CreateWebResourceResponse(new MemoryStream(shim), 200, "OK", "Content-Type: text/javascript; charset=utf-8");
            };
            core.Settings.IsStatusBarEnabled = false;
            core.Settings.AreDevToolsEnabled = true; // right-click > Inspect, for troubleshooting
            core.NewWindowRequested += (s, e) =>
            {
                e.Handled = true;
                OpenInBrowser(e.Uri);
            };
            if (ownPagesOnly)
            {
                core.NavigationStarting += (s, e) =>
                {
                    if (IsOwnPage(e.Uri)) return;
                    e.Cancel = true;
                    OpenInBrowser(e.Uri);
                };
            }
        }

        static void OpenInBrowser(string uri)
        {
            try
            {
                if (Uri.TryCreate(uri, UriKind.Absolute, out var parsed) && (parsed.Scheme == Uri.UriSchemeHttps || parsed.Scheme == Uri.UriSchemeHttp))
                    Process.Start(new ProcessStartInfo(parsed.AbsoluteUri) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Log.Error("OpenInBrowser", ex);
            }
        }
    }
}
