using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace LiquidMorph.Host
{
    /// <summary>
    /// Minimal WebView2 host for the LiquidMorph renderer.
    /// Usage: LiquidMorph.Host.exe &lt;workDir&gt; &lt;resultPath&gt;
    /// The workDir must contain index.html, renderer.js, engine/, request.json and backdrop.png.
    /// The renderer posts "RESULT:&lt;base64 png&gt;" or "ERROR:&lt;message&gt;".
    /// </summary>
    internal sealed class RenderForm : Form
    {
        private const string VirtualHost = "liquidmorph.local";
        private readonly string _workDir;
        private readonly string _resultPath;
        private readonly string _userDataOverride;
        private readonly WebView2 _web;
        private readonly Timer _timeout;
        private FileStream _video;
        private bool _finished;
        private string _error;

        public int ExitCode { get; private set; }

        public RenderForm(string workDir, string resultPath, int timeoutMs, string userDataDir = null)
        {
            _workDir = Path.GetFullPath(workDir);
            _userDataOverride = userDataDir;
            _resultPath = Path.GetFullPath(resultPath);

            Text = "LiquidMorph Renderer";
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-32000, -32000);
            Size = new Size(320, 240);
            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.None;

            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);

            _timeout = new Timer { Interval = timeoutMs };
            _timeout.Tick += (s, e) => Finish(3, "render timed out");

            Shown += async (s, e) => await InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            try
            {
                string userData = _userDataOverride;
                try { File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "last-userdata.txt"), userData ?? "<null>"); } catch { }
                if (string.IsNullOrEmpty(userData)) userData = Environment.GetEnvironmentVariable("LIQUIDMORPH_USER_DATA");
                if (string.IsNullOrEmpty(userData))
                {
                    userData = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "LiquidMorphFree", "WebView2");
                }
                try
                {
                    Directory.CreateDirectory(userData);
                    string probe = Path.Combine(userData, "probe.tmp");
                    File.WriteAllText(probe, "ok");
                    File.Delete(probe);
                }
                catch
                {
                    userData = Path.Combine(Path.GetTempPath(), "LiquidMorphFree", "WebView2");
                    Directory.CreateDirectory(userData);
                }

                CoreWebView2EnvironmentOptions options = new CoreWebView2EnvironmentOptions
                {
                    AdditionalBrowserArguments = "--enable-unsafe-webgpu --ignore-gpu-blocklist",
                };
                CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(null, userData, options);
                await _web.EnsureCoreWebView2Async(environment);

                _web.CoreWebView2.Settings.AreDevToolsEnabled = false;
                _web.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                _web.CoreWebView2.Settings.IsStatusBarEnabled = false;
                _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    VirtualHost, _workDir, CoreWebView2HostResourceAccessKind.DenyCors);
                _web.CoreWebView2.WebMessageReceived += OnMessage;

                _timeout.Start();
                _web.CoreWebView2.Navigate("https://" + VirtualHost + "/index.html");
            }
            catch (Exception ex)
            {
                Finish(3, "host init failed: " + ex.Message);
            }
        }

        private void OnMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            string message;
            try { message = e.TryGetWebMessageAsString(); }
            catch { return; }

            if (message.StartsWith("RESULT:", StringComparison.Ordinal))
            {
                try
                {
                    byte[] png = Convert.FromBase64String(message.Substring("RESULT:".Length));
                    Directory.CreateDirectory(Path.GetDirectoryName(_resultPath));
                    File.WriteAllBytes(_resultPath, png);
                    Finish(0, null);
                }
                catch (Exception ex)
                {
                    Finish(3, "write result failed: " + ex.Message);
                }
            }
            else if (message.StartsWith("VIDEO_CHUNK:", StringComparison.Ordinal))
            {
                try
                {
                    if (_video == null) _video = File.Create(_resultPath + ".part");
                    byte[] chunk = Convert.FromBase64String(message.Substring(12));
                    _video.Write(chunk, 0, chunk.Length);
                    _web.CoreWebView2.PostWebMessageAsString("ACK");
                }
                catch (Exception ex) { Finish(3, ex.Message); }
            }
            else if (message == "VIDEO_DONE" && _video != null)
            {
                try
                {
                    _video.Dispose();
                    _video = null;
                    File.Move(_resultPath + ".part", _resultPath);
                    Finish(0, null);
                }
                catch (Exception ex) { Finish(3, ex.Message); }
            }
            else if (message.StartsWith("ERROR:", StringComparison.Ordinal))
            {
                Finish(3, message.Substring("ERROR:".Length));
            }
        }

        private void Finish(int exitCode, string error)
        {
            if (_finished) return;
            _finished = true;
            if (_video != null) { _video.Dispose(); File.Delete(_resultPath + ".part"); }
            ExitCode = exitCode;
            _error = error;
            _timeout.Stop();
            if (!string.IsNullOrEmpty(error))
            {
                Console.Error.WriteLine(error);
                try { File.WriteAllText(_resultPath + ".error.txt", error); } catch { }
                try { File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "last-error.txt"), error); } catch { }
            }
            Close();
        }
    }

    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.Error.WriteLine("usage: LiquidMorph.Host.exe <workDir> <resultPath>");
                return 2;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            RenderForm form = new RenderForm(args[0], args[1], 120000, args.Length >= 3 ? args[2] : null);
            Application.Run(form);
            return form.ExitCode;
        }
    }
}



