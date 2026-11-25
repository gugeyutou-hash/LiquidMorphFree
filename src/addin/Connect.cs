using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

[assembly: AssemblyVersion("1.0.1.0")]

namespace LiquidMorph.Addin
{
    [ComVisible(true)]
    [Guid("7647A779-F679-48FF-8A74-A853D0233B01")]
    [ProgId("LiquidMorph.Free.Desktop")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object _application;

        public string GetCustomUI(string ribbonId)
        {
            Assembly assembly = typeof(Connect).Assembly;
            using (Stream stream = assembly.GetManifestResourceStream("LiquidMorph.Addin.Ribbon.xml"))
            {
                if (stream == null) return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"></customUI>";
                using (StreamReader reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        public void OnInsertGlass(IRibbonControl control)
        {
            try
            {
                dynamic app = _application;
                dynamic window = app.ActiveWindow;
                if (window == null) throw new InvalidOperationException("没有打开的演示文稿。");
                dynamic slide = window.View.Slide;
                dynamic selection = window.Selection;
                dynamic shape = SelectedShape(selection);
                if (shape == null)
                {
                    Report("请先选中一个形状，再生成液态玻璃。", null);
                    return;
                }


                GlassOptions options = GlassOptionsStore.Load();
                GlassResult result = GlassPipeline.Render(slide, shape, options);
                GlassPipeline.InsertIntoSlide(slide, result, options);
                if (options.ReplaceShape)
                {
                    try { shape.Delete(); } catch { }
                }
            }
            catch (Exception ex)
            {
                Report(ex.Message, ex);
            }
        }

        public void OnInsertCentered(IRibbonControl control)
        {
            try
            {
                dynamic app = _application;
                object slide = app.ActiveWindow.View.Slide;
                GlassOptions options = GlassOptionsStore.Load();
                GlassResult result = GlassPipeline.RenderCentered(slide, int.Parse(control.Tag), options);
                GlassPipeline.InsertIntoSlide(slide, result, options);
            }
            catch (Exception ex) { Report(ex.Message, ex); }
        }

        private static void Report(string message, Exception error)
        {
            string logPath = Environment.GetEnvironmentVariable("LIQUIDMORPH_LOG");
            if (string.IsNullOrEmpty(logPath))
            {
                try
                {
                    string dir = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "LiquidMorph");
                    System.IO.Directory.CreateDirectory(dir);
                    logPath = System.IO.Path.Combine(dir, "addin.log");
                }
                catch { logPath = null; }
            }
            if (!string.IsNullOrEmpty(logPath))
            {
                try
                {
                    System.IO.File.AppendAllText(logPath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message +
                        (error != null ? Environment.NewLine + error : string.Empty) + Environment.NewLine);
                }
                catch { }
            }

            if (Environment.GetEnvironmentVariable("LIQUIDMORPH_NOUI") == "1")
                return;

            MessageBox.Show(message, "LiquidMorph", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _application = Application;
            WatchSelection();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if(selectionTimer!=null){selectionTimer.Dispose();selectionTimer=null;}
            _application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
    }
}





