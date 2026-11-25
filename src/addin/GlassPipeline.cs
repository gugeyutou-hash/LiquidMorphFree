using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;

namespace LiquidMorph.Addin
{
    /// <summary>
    /// Renders a glass panel for a PowerPoint shape:
    ///   export slide → crop padded region → run LiquidMorph.Host → PNG
    /// </summary>
    public static class GlassPipeline
    {
        public static GlassResult Render(object slide, object shape, GlassOptions options)
        {
            dynamic s = slide;
            dynamic sh = shape;
            int kind = 5;
            if ((int)sh.Type == 1)
            {
                int autoShape = (int)sh.AutoShapeType;
                if (autoShape == 1 || autoShape == 7 || autoShape == 9) kind = autoShape;
            }
            int stored;
            if (int.TryParse((string)sh.Tags.Item("LM_KIND"), out stored)) kind = stored;
            return RenderRegion(slide, (double)sh.Left, (double)sh.Top, (double)sh.Width,
                (double)sh.Height, kind, shape, options);
        }

        public static GlassResult RenderCentered(object slide, int kind, GlassOptions options)
        {
            dynamic s = slide;
            dynamic setup = s.Parent.PageSetup;
            const double size = 170.08;
            return RenderRegion(slide, ((double)setup.SlideWidth - size) / 2,
                ((double)setup.SlideHeight - size) / 2, size, size, kind, null, options);
        }

        private static GlassResult RenderRegion(object slide, double left, double top,
            double width, double height, int kind, object source, GlassOptions options)
        {
            dynamic s = slide;
            dynamic presentation = s.Parent;
            double slideWidth = (double)presentation.PageSetup.SlideWidth;
            double slideHeight = (double)presentation.PageSetup.SlideHeight;

            double pad = options.Padding;
            pad = Math.Min(pad, Math.Min(Math.Min(left, top),
                Math.Min(slideWidth - left - width, slideHeight - top - height)));
            pad = Math.Max(0, pad);

            double scale = options.Density;
            int slidePxW = (int)Math.Round(slideWidth * scale);
            int slidePxH = (int)Math.Round(slideHeight * scale);

            string tempRoot = ResolveWorkRoot();
            string temp = Path.Combine(tempRoot, "lm-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            string slidePng = Path.Combine(temp, "slide.png");
            dynamic original = source;
            int visibility = source == null ? 0 : (int)original.Visible;
            try
            {
                if (source != null) original.Visible = 0;
                s.Export(slidePng, "PNG", slidePxW, slidePxH);
            }
            finally
            {
                if (source != null) original.Visible = visibility;
            }

            int cropX = (int)Math.Round((left - pad) * scale);
            int cropY = (int)Math.Round((top - pad) * scale);
            int cropW = (int)Math.Round((width + 2 * pad) * scale);
            int cropH = (int)Math.Round((height + 2 * pad) * scale);

            string work = Path.Combine(temp, "render");
            Directory.CreateDirectory(work);
            CopyEngineAssets(work);

            using (Bitmap slideBitmap = new Bitmap(slidePng))
            {
                cropX = Clamp(cropX, 0, slideBitmap.Width - 1);
                cropY = Clamp(cropY, 0, slideBitmap.Height - 1);
                cropW = Clamp(cropW, 1, slideBitmap.Width - cropX);
                cropH = Clamp(cropH, 1, slideBitmap.Height - cropY);

                using (Bitmap backdrop = new Bitmap(cropW, cropH, PixelFormat.Format32bppArgb))
                {
                    using (Graphics g = Graphics.FromImage(backdrop))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.DrawImage(slideBitmap, new Rectangle(0, 0, cropW, cropH),
                            new Rectangle(cropX, cropY, cropW, cropH), GraphicsUnit.Pixel);
                    }
                    backdrop.Save(Path.Combine(work, "backdrop.png"), ImageFormat.Png);
                }
            }

            double radius = Math.Min(width, height) * options.RadiusFactor;
            File.WriteAllText(Path.Combine(work, "request.json"),
                BuildRequest(width, height, pad, radius, kind, options), new UTF8Encoding(false));

            string resultPng = Path.Combine(temp, "glass.png");
            RunHost(work, resultPng, Path.Combine(tempRoot, "webview2"));

            GlassResult result = new GlassResult
            {
                Kind = kind,
                ImagePath = resultPng,
                Left = left,
                Top = top,
                Width = width,
                Height = height,
            };
            return result;
        }

        public static void InsertIntoSlide(object slide, GlassResult result, GlassOptions options)
        {
            dynamic picture = AddPicture(slide,result,options);
            picture.Select();
        }

        internal static object AddPicture(object slide, GlassResult result, GlassOptions options)
        {
            dynamic s = slide;
            const int msoFalse = 0;
            const int msoTrue = -1;
            float left = (float)result.Left;
            float top = (float)result.Top;
            float width = (float)result.Width;
            float height = (float)result.Height;
            dynamic picture = s.Shapes.AddPicture(result.ImagePath, msoFalse, msoTrue, left, top, width, height);
            picture.Tags.Add("LM_KIND", result.Kind.ToString());
            picture.Tags.Add("LM_OPTIONS", EncodeOptions(options));
            if (options.Shadow > 0)
            {
                picture.Shadow.Visible = -1;
                picture.Shadow.Transparency = (float)(1 - options.Shadow / 100);
                picture.Shadow.Blur = 7f;
            }
            picture.Tags.Add("LM_RENDER_POSE", PoseText(picture));
            return picture;
        }

        internal static string EncodeOptions(GlassOptions options)
        {
            return BitConverter.ToString(Encoding.UTF8.GetBytes(GlassOptionsStore.Json.Serialize(options))).Replace("-", "");
        }

        internal static GlassOptions DecodeOptions(string hex)
        {
            byte[] bytes=new byte[hex.Length/2];
            for(int i=0;i<bytes.Length;i++) bytes[i]=Convert.ToByte(hex.Substring(i*2,2),16);
            return GlassOptionsStore.Json.Deserialize<GlassOptions>(Encoding.UTF8.GetString(bytes));
        }

        internal static string PoseText(object shape)
        {
            dynamic sh = shape;
            return string.Format(InvariantCulture, "{0}|{1}|{2}|{3}",sh.Left,sh.Top,sh.Width,sh.Height);
        }

        private static object ParsePose(string text)
        {
            string[] values = text.Split('|');
            if (values.Length != 4) throw new InvalidOperationException("请在同一对象上设置起点和终点。");
            return new { Left=double.Parse(values[0],InvariantCulture), Top=double.Parse(values[1],InvariantCulture),
                Width=double.Parse(values[2],InvariantCulture), Height=double.Parse(values[3],InvariantCulture) };
        }

        internal static void Animate(object slide, object shape, GlassOptions o)
        {
            dynamic s=slide, sh=shape;
            var json=GlassOptionsStore.Json;
            var r=json.Deserialize<System.Collections.Generic.Dictionary<string,object>>(json.Serialize(o));
            r["A"]=ParsePose((string)sh.Tags.Item("LM_POINT_A"));
            r["B"]=ParsePose((string)sh.Tags.Item("LM_POINT_B"));
            int kind;
            if(!int.TryParse((string)sh.Tags.Item("LM_KIND"),out kind)) kind=(int)sh.Type==1?(int)sh.AutoShapeType:5;
            r["Kind"]=kind;
            r["PatternAngle"]=o.ShaderPatternAngle();
            foreach(string endpoint in new[]{"A","B"})
            {
                string encoded=sh.Tags.Item("LM_STYLE_"+endpoint);
                GlassOptions style=string.IsNullOrEmpty(encoded)?o:DecodeOptions(encoded);
                if(style.Material!=o.Material) throw new InvalidOperationException("请给起点和终点设置相同材质。");
                var appearance=json.Deserialize<System.Collections.Generic.Dictionary<string,object>>(json.Serialize(style));
                appearance["Kind"]=kind;appearance["CornerRatio"]=style.RadiusFactor;
                appearance["PatternAngle"]=style.ShaderPatternAngle();
                r["Appearance"+endpoint]=appearance;
            }
            r["Mode"]=o.Material==18?"motion":"materialMotion";
            double w=s.Parent.PageSetup.SlideWidth,h=s.Parent.PageSetup.SlideHeight;
            int px=Math.Max(2,(int)Math.Round(w/2)*2),py=Math.Max(2,(int)Math.Round(h/2)*2);
            r["PageWidth"]=w;r["PageHeight"]=h;r["PixelWidth"]=px;r["PixelHeight"]=py;
            r["CornerRatio"]=o.RadiusFactor;r["MaximumRefraction"]=1.78787875;
            r["ShadowX"]=0;r["ShadowY"]=3;r["ShadowBlur"]=7;
            const string settings=@"HKEY_CURRENT_USER\Software\LiquidMorph\Free";
            var path=new System.Collections.Generic.Dictionary<string,double>();
            foreach(string setting in new[]{"PathX","PathY","EasingX1","EasingY1","EasingX2","EasingY2"})
            {
                double value;
                if(!double.TryParse(Microsoft.Win32.Registry.GetValue(settings,setting,null) as string,
                    NumberStyles.Float,InvariantCulture,out value)) continue;
                if(setting.StartsWith("Path")) path["D"+setting.Substring(4)]=value;
                else r[setting]=value;
            }
            r["Path"]=path;
            r["Fps"]=Microsoft.Win32.Registry.GetValue(settings,"Fps",30)??30;
            r["Seconds"]=Microsoft.Win32.Registry.GetValue(settings,"Seconds",2)??2;
            string root=ResolveWorkRoot(), work=Path.Combine(root,"motion-"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(work);CopyEngineAssets(work);
            int visible=(int)sh.Visible;
            try {sh.Visible=0;s.Export(Path.Combine(work,"backdrop.png"),"PNG",px,py);}
            finally {sh.Visible=visible;}
            File.WriteAllText(Path.Combine(work,"request.json"),json.Serialize(r),new UTF8Encoding(false));
            string movie=Path.Combine(work,"motion.mp4");
            RunHost(work,movie,Path.Combine(root,"webview2"));
            dynamic video=s.Shapes.AddMediaObject2(movie,0,-1,0f,0f,(float)w,(float)h);
            video.Select();
        }

        private static string ResolveWorkRoot()
        {
            string fromEnv = Environment.GetEnvironmentVariable("LIQUIDMORPH_TEMP");
            if (!string.IsNullOrEmpty(fromEnv) && TryCreate(fromEnv)) return fromEnv;

            string local = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LiquidMorph", "work");
            if (TryCreate(local)) return local;

            // Portable installations can keep the exported backdrop, browser profile and
            // result together in a location accessible to the separate renderer process.
            string beside = Path.Combine(
                Path.GetDirectoryName(typeof(GlassPipeline).Assembly.Location), "work");
            if (TryCreate(beside)) return beside;

            if (TryCreate(Path.GetTempPath())) return Path.GetTempPath();

            throw new DirectoryNotFoundException(
                "No writable working directory. Set LIQUIDMORPH_TEMP to a writable folder.");
        }

        private static bool TryCreate(string path)
        {
            try
            {
                Directory.CreateDirectory(path);
                string probe = Path.Combine(path, "probe-" + Guid.NewGuid().ToString("N") + ".tmp");
                File.WriteAllText(probe, "ok");
                File.Delete(probe);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void RunHost(string work, string resultPng, string userData)
        {
            string host = LocateHost();
            ProcessStartInfo psi = new ProcessStartInfo(host, Quote(work) + " " + Quote(resultPng) + " " + Quote(userData))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = Path.GetDirectoryName(host),
            };
            psi.RedirectStandardError = true;
            psi.EnvironmentVariables["LIQUIDMORPH_USER_DATA"] = userData;
            Log("host=" + host + " args=" + psi.Arguments);
            using (Process process = Process.Start(psi))
            {
                var errorOutput = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(180000))
                {
                    try { process.Kill(); } catch { }
                    throw new TimeoutException("LiquidMorph renderer timed out.");
                }
                if (process.ExitCode != 0 || !File.Exists(resultPng))
                {
                    string detail = string.Empty;
                    try
                    {
                        string errorFile = resultPng + ".error.txt";
                        if (File.Exists(errorFile)) detail = ": " + File.ReadAllText(errorFile).Trim();
                        else if (!string.IsNullOrWhiteSpace(errorOutput.Result)) detail = ": " + errorOutput.Result.Trim();
                    }
                    catch { }
                    throw new InvalidOperationException(
                        "LiquidMorph renderer failed (exit " + process.ExitCode + ") in " + work + detail + ".");
                }
            }
        }

        private static string LocateHost()
        {
            string fromEnv = Environment.GetEnvironmentVariable("LIQUIDMORPH_HOST");
            if (!string.IsNullOrEmpty(fromEnv) && File.Exists(fromEnv)) return fromEnv;

            string dir = Path.GetDirectoryName(typeof(GlassPipeline).Assembly.Location);
            string[] candidates =
            {
                Path.Combine(dir, "LiquidMorph.Host.exe"),
                Path.Combine(dir, "host", "LiquidMorph.Host.exe"),
                Path.Combine(dir, "..", "host", "LiquidMorph.Host.exe"),
            };
            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate)) return Path.GetFullPath(candidate);
            }
            throw new FileNotFoundException("LiquidMorph.Host.exe not found. Set LIQUIDMORPH_HOST.");
        }

        private static void CopyEngineAssets(string work)
        {
            string engine = Environment.GetEnvironmentVariable("LIQUIDMORPH_ENGINE");
            if (string.IsNullOrEmpty(engine) || !Directory.Exists(engine))
            {
                string dir = Path.GetDirectoryName(typeof(GlassPipeline).Assembly.Location);
                string[] candidates =
                {
                    Path.Combine(dir, "engine"),
                    Path.Combine(dir, "..", "engine"),
                    Path.Combine(dir, "..", "..", "src", "engine"),
                };
                foreach (string candidate in candidates)
                {
                    if (Directory.Exists(candidate)) { engine = candidate; break; }
                }
            }
            if (string.IsNullOrEmpty(engine) || !Directory.Exists(engine))
                throw new DirectoryNotFoundException("engine assets not found. Set LIQUIDMORPH_ENGINE.");

            foreach (string file in Directory.GetFiles(engine, "*", SearchOption.AllDirectories))
            {
                string relative = file.Substring(engine.Length).TrimStart('\\', '/');
                string target = Path.Combine(work, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                File.Copy(file, target, true);
            }
        }

        private static string BuildRequest(double width, double height, double pad, double radius, int kind, GlassOptions o)
        {
            string pixelW = ((int)Math.Round(width * o.Density)).ToString(InvariantCulture);
            string pixelH = ((int)Math.Round(height * o.Density)).ToString(InvariantCulture);
            StringBuilder sb = new StringBuilder();
            sb.Append("{");
            AppendNumber(sb, "Width", width); sb.Append(",");
            AppendNumber(sb, "Height", height); sb.Append(",");
            AppendNumber(sb, "Padding", pad); sb.Append(",");
            AppendNumber(sb, "Kind", kind); sb.Append(",");
            AppendNumber(sb, "Radius", radius); sb.Append(",");
            AppendNumber(sb, "Density", o.Density); sb.Append(",");
            sb.Append("\"PixelWidth\":").Append(pixelW).Append(",");
            sb.Append("\"PixelHeight\":").Append(pixelH).Append(",");
            AppendNumber(sb, "Strength", o.Strength); sb.Append(",");
            AppendNumber(sb, "GlowWidth", o.GlowWidth); sb.Append(",");
            AppendNumber(sb, "Spread", o.Spread); sb.Append(",");
            AppendNumber(sb, "LightAngle", o.LightAngle); sb.Append(",");
            sb.Append("\"EdgeType\":").Append(o.EdgeType).Append(",");
            AppendNumber(sb, "Refraction", o.Refraction == 0 ? 0 : 1); sb.Append(",");
            AppendNumber(sb, "RefractionStrength", o.RefractionStrength); sb.Append(",");
            AppendNumber(sb, "MaximumRefraction", 1.78787875); sb.Append(",");
            AppendNumber(sb, "Dispersion", o.Dispersion == 0 ? 0 : 1); sb.Append(",");
            AppendNumber(sb, "DispersionStrength", o.DispersionStrength); sb.Append(",");
            AppendNumber(sb, "Red", o.Red); sb.Append(",");
            AppendNumber(sb, "Green", o.Green); sb.Append(",");
            AppendNumber(sb, "Blue", o.Blue); sb.Append(",");
            AppendNumber(sb, "Material", o.Material); sb.Append(",");
            AppendNumber(sb, "Magnification", o.Magnification); sb.Append(",");
            AppendNumber(sb, "Blur", o.Blur); sb.Append(",");
            AppendNumber(sb, "PatternScale", o.PatternScale); sb.Append(",");
            AppendNumber(sb, "PatternDepth", o.PatternDepth); sb.Append(",");
            AppendNumber(sb, "PatternAngle", o.ShaderPatternAngle()); sb.Append(",");
            AppendNumber(sb, "Tint", o.Tint);
            sb.Append("}");
            return sb.ToString();
        }

        private static CultureInfo InvariantCulture = CultureInfo.InvariantCulture;

        private static void AppendNumber(StringBuilder sb, string name, double value)
        {
            sb.Append('"').Append(name).Append("\":")
              .Append(Math.Round(value, 4).ToString(InvariantCulture));
        }

        private static void Log(string message)
        {
            string path = Environment.GetEnvironmentVariable("LIQUIDMORPH_LOG");
            if (string.IsNullOrEmpty(path)) return;
            try { File.AppendAllText(path, DateTime.Now.ToString("HH:mm:ss.fff") + " " + message + Environment.NewLine); } catch { }
        }

        private static string Quote(string value)
        {
            return "\"" + value + "\"";
        }

        private static int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }
}








