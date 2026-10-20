using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Claude Mail Assistant Setup")]
[assembly: AssemblyProduct("Claude Mail Assistant")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace ClaudeMailAssistant
{
    // Installs the Outlook add-in for the current Windows user, without administrator rights:
    //   ClaudeMailAssistantSetup.exe              install or update
    //   ClaudeMailAssistantSetup.exe /uninstall   remove it (Settings > Apps runs this)
    //   add /quiet to skip the dialogs
    // The add-in's files are embedded in this exe as resources named "payload/<relative path>".
    static class Setup
    {
        const string PayloadPrefix = "payload/";
        const string AddinKey = @"Software\Microsoft\Office\Outlook\Addins\" + Ids.ConnectProgId;
        const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeMailAssistant";
        const string DoNotDisableKey = @"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList";
        const string ManagedCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";
        const string ControlCategory = "{40FC6ED4-2438-11CF-A3DB-080036F12502}";

        static readonly string LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        static readonly string InstallDir = Path.Combine(LocalAppData, "Programs", "ClaudeMailAssistant");
        static readonly string DataDir = Path.Combine(LocalAppData, "ClaudeMailAssistant");
        static readonly string AddinDll = Path.Combine(InstallDir, "ClaudeMailAssistant.dll");
        static readonly string SetupExe = Path.Combine(InstallDir, "ClaudeMailAssistantSetup.exe");

        static bool quiet;

        [STAThread]
        static int Main(string[] args)
        {
            quiet = args.Any(a => a.Equals("/quiet", StringComparison.OrdinalIgnoreCase));
            bool uninstall = args.Any(a => a.Equals("/uninstall", StringComparison.OrdinalIgnoreCase));
            Application.EnableVisualStyles();
            try
            {
                if (uninstall)
                {
                    if (!Ask("Remove Claude Mail Assistant from Outlook?\n\nThis also deletes its settings, including your saved Claude API key.")) return 1;
                    WaitForOutlookToClose();
                    Uninstall();
                    Tell("Claude Mail Assistant was removed from Outlook.");
                }
                else
                {
                    if (!Ask($"Install Claude Mail Assistant {Ids.Version} into Outlook?\n\nIt's installed for your Windows account only and doesn't need administrator rights.")) return 1;
                    if (!WebView2Installed() && !Ask("Claude needs the Microsoft Edge WebView2 Runtime, which isn't on this PC. Get it from https://go.microsoft.com/fwlink/p/?LinkId=2124703.\n\nInstall Claude anyway?")) return 1;
                    WaitForOutlookToClose();
                    Install();
                    Tell("Claude Mail Assistant is installed.\n\nStart Outlook, select an email and click Claude on the Home tab. Then open Settings in the Claude pane and paste your Claude API key.");
                }
                return 0;
            }
            catch (OperationCanceledException)
            {
                return 1;
            }
            catch (Exception ex)
            {
                if (!quiet) MessageBox.Show("Setup failed: " + ex.Message, Ids.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }
        }

        static void Install()
        {
            var assembly = typeof(Setup).Assembly;
            var payload = assembly.GetManifestResourceNames().Where(n => n.StartsWith(PayloadPrefix, StringComparison.Ordinal)).ToArray();
            if (payload.Length == 0) throw new InvalidOperationException("This setup file is incomplete. Download it again.");

            DeleteDirectory(Path.Combine(InstallDir, "public")); // drop files a newer version no longer ships
            foreach (var name in payload)
            {
                var target = Path.Combine(InstallDir, name.Substring(PayloadPrefix.Length).Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                using (var source = assembly.GetManifestResourceStream(name))
                using (var file = File.Create(target))
                    source.CopyTo(file);
            }
            var self = Path.GetFullPath(assembly.Location);
            if (!self.Equals(Path.GetFullPath(SetupExe), StringComparison.OrdinalIgnoreCase)) File.Copy(self, SetupExe, true);

            // COM registration, the per-user equivalent of "regasm /codebase". Both registry views, so it
            // works with 32-bit and 64-bit Outlook.
            var assemblyName = AssemblyName.GetAssemblyName(AddinDll);
            var codeBase = new Uri(AddinDll).AbsoluteUri;
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    RegisterClass(root, Ids.ConnectClsid, Ids.ConnectProgId, assemblyName, codeBase, isControl: false);
                    RegisterClass(root, Ids.PaneClsid, Ids.PaneProgId, assemblyName, codeBase, isControl: true);
                }
            }

            using (var key = Registry.CurrentUser.CreateSubKey(AddinKey))
            {
                key.SetValue("FriendlyName", Ids.ProductName);
                key.SetValue("Description", "Summarize, draft replies, triage, spam-check and search your mail with Claude.");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord); // load at startup
            }
            using (var key = Registry.CurrentUser.CreateSubKey(DoNotDisableKey))
                key.SetValue(Ids.ConnectProgId, 1, RegistryValueKind.DWord); // don't auto-disable after a slow start

            using (var key = Registry.CurrentUser.CreateSubKey(UninstallKey))
            {
                key.SetValue("DisplayName", Ids.ProductName);
                key.SetValue("DisplayVersion", Ids.Version);
                key.SetValue("Publisher", Ids.ProductName);
                key.SetValue("DisplayIcon", SetupExe);
                key.SetValue("InstallLocation", InstallDir);
                key.SetValue("UninstallString", $"\"{SetupExe}\" /uninstall");
                key.SetValue("QuietUninstallString", $"\"{SetupExe}\" /uninstall /quiet");
                key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                key.SetValue("EstimatedSize", (int)(DirectorySize(InstallDir) / 1024), RegistryValueKind.DWord);
            }
        }

        static void RegisterClass(RegistryKey root, string clsid, string progId, AssemblyName assembly, string codeBase, bool isControl)
        {
            var id = "{" + clsid + "}";
            using (var key = root.CreateSubKey(@"Software\Classes\" + progId))
            {
                key.SetValue("", progId);
                using (var sub = key.CreateSubKey("CLSID")) sub.SetValue("", id);
            }
            using (var key = root.CreateSubKey(@"Software\Classes\CLSID\" + id))
            {
                key.SetValue("", progId);
                using (var sub = key.CreateSubKey("ProgId")) sub.SetValue("", progId);
                using (var server = key.CreateSubKey("InprocServer32"))
                {
                    server.SetValue("", "mscoree.dll");
                    server.SetValue("ThreadingModel", "Both");
                    WriteClrValues(server, progId, assembly, codeBase);
                    using (var version = server.CreateSubKey(assembly.Version.ToString())) WriteClrValues(version, progId, assembly, codeBase);
                }
                key.CreateSubKey(@"Implemented Categories\" + ManagedCategory).Dispose();
                if (isControl)
                {
                    key.CreateSubKey("Control").Dispose();
                    key.CreateSubKey(@"Implemented Categories\" + ControlCategory).Dispose();
                }
            }
        }

        static void WriteClrValues(RegistryKey key, string className, AssemblyName assembly, string codeBase)
        {
            key.SetValue("Class", className); // the ProgID is the class's full name
            key.SetValue("Assembly", assembly.FullName);
            key.SetValue("RuntimeVersion", "v4.0.30319");
            key.SetValue("CodeBase", codeBase);
        }

        static void Uninstall()
        {
            Registry.CurrentUser.DeleteSubKeyTree(AddinKey, false);
            Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false);
            using (var key = Registry.CurrentUser.OpenSubKey(DoNotDisableKey, true)) key?.DeleteValue(Ids.ConnectProgId, false);
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    foreach (var (clsid, progId) in new[] { (Ids.ConnectClsid, Ids.ConnectProgId), (Ids.PaneClsid, Ids.PaneProgId) })
                    {
                        root.DeleteSubKeyTree(@"Software\Classes\CLSID\{" + clsid + "}", false);
                        root.DeleteSubKeyTree(@"Software\Classes\" + progId, false);
                    }
                }
            }
            DeleteDirectory(DataDir);

            var self = Path.GetFullPath(typeof(Setup).Assembly.Location);
            if (self.StartsWith(InstallDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                // This exe can't delete its own folder while running; let cmd do it once we've exited.
                Process.Start(new ProcessStartInfo("cmd.exe", $"/c ping -n 3 127.0.0.1 >nul & rmdir /s /q \"{InstallDir}\"")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WorkingDirectory = Path.GetTempPath(),
                });
            }
            else
            {
                DeleteDirectory(InstallDir);
            }
        }

        static void WaitForOutlookToClose()
        {
            while (Process.GetProcessesByName("OUTLOOK").Length > 0)
            {
                if (quiet) throw new InvalidOperationException("Close Outlook first.");
                if (MessageBox.Show("Please close Outlook, then choose OK.", Ids.ProductName, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK)
                    throw new OperationCanceledException();
                Thread.Sleep(1500); // Outlook takes a moment to exit after its window closes
            }
        }

        static bool WebView2Installed()
        {
            const string client = @"Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";
            var keys = new[] { @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\" + client, @"HKEY_LOCAL_MACHINE\SOFTWARE\" + client, @"HKEY_CURRENT_USER\Software\" + client };
            return keys.Any(k => Registry.GetValue(k, "pv", null) is string pv && pv.Length > 0 && pv != "0.0.0.0");
        }

        static long DirectorySize(string dir) =>
            Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories).Sum(f => new FileInfo(f).Length);

        static void DeleteDirectory(string dir)
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }

        static bool Ask(string text) =>
            quiet || MessageBox.Show(text, Ids.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

        static void Tell(string text)
        {
            if (!quiet) MessageBox.Show(text, Ids.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
