using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Win32;
using FileDialog = System.Windows.Forms.FileDialog;

namespace LiquidMorph.Addin
{
    public partial class Connect
    {
        private IRibbonUI ribbon;
        private Timer selectionTimer;
        private string lastSelection;

        private readonly Dictionary<int,string> baseline=new Dictionary<int,string>();
        private readonly Dictionary<int,string> pending=new Dictionary<int,string>();
        private object watchedSlide;

        private void CheckMovedGlass(dynamic slide,dynamic selection)
        {
            if(!object.Equals(watchedSlide,(object)slide))
            {
                watchedSlide=slide;
                baseline.Clear();
                pending.Clear();
            }
            if(Control.MouseButtons!=MouseButtons.None)return;
            dynamic shapes=slide.Shapes;
            int count=shapes.Count;
            for(int i=1;i<=count;i++)
            {
                dynamic shape=shapes[i];
                if((int)shape.Type!=13 || string.IsNullOrEmpty((string)shape.Tags.Item("LM_KIND")))continue;
                int id=shape.Id;
                string pose=GlassPipeline.PoseText(shape);
                string rendered=shape.Tags.Item("LM_RENDER_POSE");
                if(string.IsNullOrEmpty(rendered) && !baseline.TryGetValue(id,out rendered))
                    baseline[id]=rendered=pose;
                if(rendered==pose){pending.Remove(id);continue;}
                string previous;
                if(!pending.TryGetValue(id,out previous) || previous!=pose){pending[id]=pose;continue;}
                // A repeated failure at this pose should not open a dialog every tick.
                if(baseline.ContainsKey(id) && baseline[id]=="failed:"+pose)continue;
                try
                {
                    RefreshMovedGlass(slide,shape,selection);
                    pending.Remove(id);
                    baseline.Remove(id);
                }
                catch(Exception ex)
                {
                    baseline[id]="failed:"+pose;
                    Report("移动后的玻璃未能刷新："+ex.Message,ex);
                }
                return;
            }
        }

        private static void RefreshMovedGlass(dynamic slide,dynamic shape,dynamic selection)
        {
            bool selected=false;
            if((int)selection.Type==2)
            {
                dynamic range=selection.ShapeRange;
                for(int i=1;i<=(int)range.Count;i++)
                    if((int)range[i].Id==(int)shape.Id){selected=true;break;}
            }
            int z=shape.ZOrderPosition;
            string name=shape.Name;
            GlassOptions options=ReadOptions(shape);
            GlassResult result=GlassPipeline.Render(slide,shape,options);
            dynamic replacement=GlassPipeline.AddPicture(slide,result,options);
            try
            {
                replacement.AlternativeText=shape.AlternativeText;
                replacement.Rotation=shape.Rotation;
                dynamic tags=shape.Tags;
                for(int i=1;i<=(int)tags.Count;i++)
                {
                    string key=tags.Name(i);
                    if(key!="LM_RENDER_POSE")replacement.Tags.Add(key,tags.Item(key));
                }
                while((int)replacement.ZOrderPosition>z+1)replacement.ZOrder(3);
                if(selected)replacement.Select(0);
                shape.Delete();
            }
            catch {replacement.Delete();throw;}
            replacement.Name=name;
        }

        private void WatchSelection()
        {
            selectionTimer=new Timer {Interval=500};
            selectionTimer.Tick += delegate {
                string selection="";
                try
                {
                    dynamic app=_application, window=app.ActiveWindow, current=window.Selection;
                    selectionTimer.Stop();
                    CheckMovedGlass(window.View.Slide,current);
                    dynamic shape=SelectedShape(current);
                    selection=window.View.Slide.SlideID+"/"+current.Type+"/"+(shape==null ? 0 : (int)shape.Id);
                }
                catch { }
                finally {if(selectionTimer!=null)selectionTimer.Start();}
                if(selection!=lastSelection && ribbon!=null){lastSelection=selection;ribbon.Invalidate();}
            };
            selectionTimer.Start();
        }

        private static object SelectedShape(dynamic selection)
        {
            int type=(int)selection.Type;
            if(type!=2 && type!=3)return null;
            if(type==3)
            {
                try {return selection.TextRange.Parent.Parent;}
                catch { }
            }
            dynamic range=selection.ShapeRange;
            return (int)range.Count==1 ? range[1] : null;
        }

        public bool ShapeSelected(IRibbonControl control)
        {
            try {dynamic app=_application, selection=app.ActiveWindow.Selection;return SelectedShape(selection)!=null;}
            catch{return false;}
        }

        public bool GlassCommandEnabled(IRibbonControl control)
        {
            if(!ShapeSelected(control)) return false;
            try
            {
                dynamic app=_application, shape=SelectedShape(app.ActiveWindow.Selection);
                if(string.IsNullOrEmpty((string)shape.Tags.Item("LM_KIND"))) return false;
                bool start=!string.IsNullOrEmpty((string)shape.Tags.Item("LM_POINT_A"));
                bool end=!string.IsNullOrEmpty((string)shape.Tags.Item("LM_POINT_B"));
                switch(control.Tag)
                {
                    case "animate": return start && end;
                    case "clearA": return start;
                    case "clearB": return end;
                    default: return true;
                }
            }
            catch{return false;}
        }

        public bool SlideAvailable(IRibbonControl control)
        {
            try {dynamic app=_application; return app.ActiveWindow.View.Slide!=null;}
            catch{return false;}
        }

        public void ToggleChoice(IRibbonControl control,bool pressed){RibbonAction(control);}

        public bool ChoicePressed(IRibbonControl control)
        {
            string key=control.Tag;
            if(key.EndsWith("Theme"))return LineTheme==(key=="lineTheme");
            try
            {
                dynamic app=_application,selection=app.ActiveWindow.Selection;
                GlassOptions o=ReadOptions(SelectedShape(selection));
                return key.StartsWith("mat") ? key=="mat"+o.Material : key=="edge"+o.EdgeType;
            }
            catch{return false;}
        }

        private const string SettingsKey = @"HKEY_CURRENT_USER\Software\LiquidMorph\Free";
        private static bool LineTheme { get { return Convert.ToInt32(Registry.GetValue(SettingsKey,"LineTheme",0)) == 1; } }
        public void RibbonLoaded(IRibbonUI value) { ribbon = value; }

        public bool PatternVisible(IRibbonControl control)
        {
            try
            {
                dynamic app=_application, selection=app.ActiveWindow.Selection;
                GlassOptions o=ReadOptions(SelectedShape(selection));
                return o.Material>=4 && o.Material<18;
            }
            catch {return false;}
        }

        private class Pictures : AxHost
        {
            private Pictures() : base("") { }
            public static object ConvertIcon(Image image) { return GetIPictureDispFromPicture(image); }
        }

        public object RibbonImage(IRibbonControl control)
        {
            string key = control.Tag ?? "";
            using (Stream stream=typeof(Connect).Assembly.GetManifestResourceStream("Icons."+key+".png"))
            {
                if(stream==null)return null;
                using(Bitmap icon=new Bitmap(stream))
                {
                    if(LineTheme)
                    {
                        using(Bitmap line=new Bitmap(icon.Width,icon.Height))
                        {
                            for(int y=1;y<icon.Height-1;y++)
                                for(int x=1;x<icon.Width-1;x++)
                                {
                                    Color pixel=icon.GetPixel(x,y);
                                    if(pixel.A>60 && (icon.GetPixel(x-1,y).A<100 || icon.GetPixel(x+1,y).A<100 ||
                                        icon.GetPixel(x,y-1).A<100 || icon.GetPixel(x,y+1).A<100 ||
                                        Math.Abs(pixel.GetBrightness()-icon.GetPixel(x,y-1).GetBrightness())>.13))
                                        line.SetPixel(x,y,Color.DimGray);
                                }
                            return Pictures.ConvertIcon(line);
                        }
                    }
                    return Pictures.ConvertIcon(icon);
                }
            }
        }

        private static GlassOptions ReadOptions(dynamic shape)
        {
            if (shape != null)
            {
                string hex = shape.Tags.Item("LM_OPTIONS");
                if (!string.IsNullOrEmpty(hex))
                {
                    return GlassPipeline.DecodeOptions(hex);
                }
            }
            return GlassOptionsStore.Load();
        }

        private static decimal? Number(string title,double value,decimal min,decimal max,decimal step)
        {
            using (Form form = new Form {Text=title,ClientSize=new Size(310,137),StartPosition=FormStartPosition.CenterScreen,
                FormBorderStyle=FormBorderStyle.FixedDialog, MaximizeBox=false,MinimizeBox=false})
            {
                NumericUpDown input = new NumericUpDown {Left=194,Top=12,Width=100,Minimum=min,Maximum=max,
                    DecimalPlaces=step>=1?0:2,Increment=step,Value=Math.Min(max,Math.Max(min,(decimal)value))};
                Button ok=new Button {Text="确定",Left=133,Top=97,DialogResult=DialogResult.OK};
                Button cancel=new Button {Text="取消",Left=219,Top=97,DialogResult=DialogResult.Cancel};
                Label label=new Label {Text=title,Left=16,Top=15,Width=175};
                TrackBar slider=new TrackBar {Left=9,Top=45,Width=288,Minimum=0,Maximum=100,TickFrequency=10};
                slider.Value=(int)Math.Round((input.Value-min)*100/(max-min));
                slider.Scroll+=delegate {input.Value=decimal.Round(min+slider.Value*(max-min)/100,input.DecimalPlaces);};
                input.ValueChanged+=delegate {slider.Value=(int)Math.Round((input.Value-min)*100/(max-min));};
                form.Controls.AddRange(new Control[]{label,slider,input,ok,cancel});form.AcceptButton=ok;form.CancelButton=cancel;
                return form.ShowDialog()==DialogResult.OK?(decimal?)input.Value:null;
            }
        }

        public void RibbonAction(IRibbonControl control)
        {
            try
            {
                string key=control.Tag;
                if (key=="about") {MessageBox.Show("LiquidMorph Free · 免费版","产品授权");return;}
                if (key.EndsWith("Theme"))
                {
                    Registry.SetValue(SettingsKey,"LineTheme",key=="lineTheme"?1:0);
                    if(ribbon!=null) ribbon.Invalidate();
                    return;
                }
                dynamic app=_application;
                dynamic window=app.ActiveWindow, slide=window.View.Slide;
                dynamic shape=SelectedShape(window.Selection);
                GlassOptions o=ReadOptions(shape);
                if(key=="PathX" || key=="PathY" || key.StartsWith("Easing"))
                {
                    bool path=key.StartsWith("Path");
                    double fallback=key=="EasingX1"?.65:key=="EasingX2"?.35:key=="EasingY2"?1:0;
                    string saved=Registry.GetValue(SettingsKey,key,null) as string;
                    double previous;
                    if(!double.TryParse(saved,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out previous))
                        previous=fallback;
                    decimal? v=Number(path?"路径偏移（磅）":"缓动曲线控制点",previous,path?-1000:0,path?1000:1,path?1:.01m);
                    if(v.HasValue)Registry.SetValue(SettingsKey,key,v.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    return;
                }
                if(key=="Seconds" || key=="Fps")
                {
                    int current=Convert.ToInt32(Registry.GetValue(SettingsKey,key,key=="Fps"?30:2));
                    decimal? value=Number(key=="Fps"?"动画帧率":"动画时长（秒）",current,1,key=="Fps"?60:30,1);
                    if(value.HasValue) Registry.SetValue(SettingsKey,key,(int)value.Value);
                    return;
                }
                if(key=="A" || key=="B" || key=="clearA" || key=="clearB" || key=="animate")
                {
                    if(shape==null) throw new InvalidOperationException("请先选择动画对象。");
                    if(key=="animate") GlassPipeline.Animate(slide,shape,o);
                    else if(key.StartsWith("clear")) shape.Tags.Delete("LM_POINT_"+key.Substring(5));
                    else {
                        shape.Tags.Add("LM_POINT_"+key,GlassPipeline.PoseText(shape));
                        shape.Tags.Add("LM_STYLE_"+key,GlassPipeline.EncodeOptions(o));
                    }
                    if(ribbon!=null) ribbon.Invalidate();
                    return;
                }
                if(key=="savePreset" || key=="loadPreset")
                {
                    using(FileDialog dialog=key=="savePreset"?(FileDialog)new SaveFileDialog():new OpenFileDialog())
                    {
                        dialog.Filter="玻璃材质预设 (*.json)|*.json";
                        if(dialog.ShowDialog()!=DialogResult.OK) return;
                        if(key=="savePreset"){File.WriteAllText(dialog.FileName,GlassOptionsStore.Json.Serialize(o));return;}
                        o=GlassOptionsStore.Json.Deserialize<GlassOptions>(File.ReadAllText(dialog.FileName));
                        if(o==null) throw new InvalidDataException("预设为空。");
                    }
                }
                else if(key=="reset") o=new GlassOptions();
                else if(key.StartsWith("mat")) o.ApplyMaterial(int.Parse(key.Substring(3)));
                else if(key.StartsWith("edge")) o.EdgeType=int.Parse(key.Substring(4));
                else if(key=="Color")
                {
                    using(ColorDialog dlg=new ColorDialog {FullOpen=true,Color=Color.FromArgb(o.Red,o.Green,o.Blue)})
                    {
                        if(dlg.ShowDialog()!=DialogResult.OK)return;
                        o.Red=dlg.Color.R;o.Green=dlg.Color.G;o.Blue=dlg.Color.B;
                        if(o.Tint==0)o.Tint=0.1;
                    }
                }
                else if(key!="refresh")
                {
                    var field=typeof(GlassOptions).GetField(key);
                    decimal min=key=="PatternScale"?.25m:key=="GlowWidth"?.1m:key=="Spread"?1:0;
                    decimal max=100;
                    switch(key)
                    {
                        case "LightAngle": max=360;break;
                        case "PatternAngle":
                            if(o.Material==6){min=.25m;max=4;}
                            else if(o.Material==16)max=3;
                            else max=o.Material==5?100:180;
                            break;
                        case "Tint": max=1;break;
                        case "PatternScale": max=4;break;
                        case "PatternDepth": case "DispersionStrength": max=3;break;
                        case "Strength": max=10;break;
                        case "GlowWidth": max=24;break;
                        case "Spread": max=30;break;
                        case "RefractionStrength": max=1.78787875m;break;
                    }
                    bool optical=key=="RefractionStrength" || key=="DispersionStrength";
                    bool enabled=key=="RefractionStrength"?o.Refraction!=0:o.Dispersion!=0;
                    bool percent=key=="Tint";
                    decimal value=Convert.ToDecimal(field.GetValue(o));
                    if(percent){value*=100;max=100;}
                    decimal step=key=="Blur"?.1m:max>=100?1m:.01m;
                    ShowParameter(ParameterTitle(key),value,min,max,step,optical,enabled,
                        delegate(decimal next,bool on){
                            field.SetValue(o,(double)(percent?next/100:next));
                            if(key=="RefractionStrength")o.Refraction=on?1:0;
                            if(key=="DispersionStrength")o.Dispersion=on?1:0;
                            shape=ApplyOptions(window,slide,shape,o);
                        });
                    return;
                }
                ApplyOptions(window,slide,shape,o);
            }
            catch(Exception ex){Report(ex.Message,ex);}
        }

        private object ApplyOptions(dynamic window,dynamic slide,dynamic shape,GlassOptions o)
        {
                if(shape!=null)
                {
                    int position=(int)shape.ZOrderPosition;
                    string[] tags={"LM_POINT_A","LM_POINT_B","LM_STYLE_A","LM_STYLE_B"};
                    var metadata=new Dictionary<string,string>();
                    foreach(string tag in tags) metadata[tag]=shape.Tags.Item(tag);
                    GlassResult result=GlassPipeline.Render(slide,shape,o);
                    GlassPipeline.InsertIntoSlide(slide,result,o);
                    dynamic replacement=window.Selection.ShapeRange[1];
                    foreach(var entry in metadata)
                        if(!string.IsNullOrEmpty(entry.Value)) replacement.Tags.Add(entry.Key,entry.Value);
                    while((int)replacement.ZOrderPosition>position+1) replacement.ZOrder(3);
                    shape.Delete();
                    shape=replacement;
                }
                GlassOptionsStore.Save(o);
                if(ribbon!=null) ribbon.Invalidate();
            return shape;
        }

        private ToolStripDropDown parameterPopup;

        private void ShowParameter(string title,decimal initial,decimal min,decimal max,decimal step,
            bool optical,bool enabled,Action<decimal,bool> apply)
        {
            if(parameterPopup!=null) parameterPopup.Close();
            int offset=optical?60:0;
            var panel=new Panel {BackColor=Color.White,Size=new Size(292+offset,48)};
            var toggle=new CheckBox {Text="启用",AccessibleName=title+"开关",Bounds=new Rectangle(7,9,52,30),
                Visible=optical,Checked=enabled};
            var input=new NumericUpDown {AccessibleName=title,ForeColor=Color.FromArgb(12,119,190),Bounds=new Rectangle(210+offset,12,72,30),
                DecimalPlaces=step>=1?0:step==.1m?1:2,Minimum=min,Maximum=max,Increment=step,
                Value=Math.Max(min,Math.Min(max,initial))};
            var slider=new TrackBar {AccessibleName=title,Bounds=new Rectangle(6+offset,8,200,32),
                AutoSize=false,TickStyle=TickStyle.None,Maximum=(int)Math.Round((max-min)/step)};
            slider.Value=(int)Math.Round((input.Value-min)/step);
            input.Enabled=slider.Enabled=!optical||enabled;
            panel.Controls.AddRange(new Control[]{toggle,slider,input});
            var popup=new ToolStripDropDown {AutoSize=false,Padding=Padding.Empty,Size=panel.Size};
            var host=new ToolStripControlHost(panel){AutoSize=false,Margin=Padding.Empty,Padding=Padding.Empty,Size=panel.Size};
            popup.Items.Add(host);
            float scale;
            using(Graphics graphics=panel.CreateGraphics()) scale=graphics.DpiX/96f;
            if(scale!=1f) panel.Scale(new SizeF(scale,scale));
            host.Size=popup.Size=panel.Size;
            int diameter=(int)Math.Round(26*scale);
            using(var outline=new GraphicsPath())
            {
                int right=panel.Width-diameter, bottom=panel.Height-diameter;
                outline.AddArc(0,0,diameter,diameter,180,90);
                outline.AddArc(right,0,diameter,diameter,270,90);
                outline.AddArc(right,bottom,diameter,diameter,0,90);
                outline.AddArc(0,bottom,diameter,diameter,90,90);
                outline.CloseFigure();
                popup.Region=new Region(outline);
            }
            decimal applied=input.Value;
            bool appliedEnabled=toggle.Checked;
            Action commit=delegate {
                if(input.Value==applied && toggle.Checked==appliedEnabled)return;
                try {apply(input.Value,toggle.Checked);applied=input.Value;appliedEnabled=toggle.Checked;}
                catch(Exception ex){popup.Close();Report(ex.Message,ex);}
            };
            slider.Scroll+=delegate {input.Value=Math.Min(max,min+slider.Value*step);};
            slider.MouseUp+=delegate {commit();};
            slider.KeyUp+=delegate {commit();};
            input.ValueChanged+=delegate {slider.Value=(int)Math.Round((input.Value-min)/step);};
            input.KeyDown+=delegate(object sender,KeyEventArgs e){
                if(e.KeyCode==Keys.Enter){commit();e.SuppressKeyPress=true;}
                if(e.KeyCode==Keys.Escape){popup.Close();e.SuppressKeyPress=true;}
            };
            toggle.CheckedChanged+=delegate {input.Enabled=slider.Enabled=toggle.Checked;commit();};
            popup.Closed+=delegate {
                if(parameterPopup==popup) parameterPopup=null;
                if(popup.IsHandleCreated) popup.BeginInvoke(new Action(popup.Dispose));
            };
            parameterPopup=popup;
            var anchor=Cursor.Position;
            Rectangle screen=Screen.FromPoint(anchor).WorkingArea;
            int x=Math.Max(screen.Left,Math.Min(screen.Right-panel.Width,anchor.X-panel.Width/2));
            int y=anchor.Y+12+panel.Height<=screen.Bottom?anchor.Y+12:anchor.Y-panel.Height-8;
            popup.Show(new Point(x,y));
        }

        private static string ParameterTitle(string key)
        {
            var names = new Dictionary<string,string> {{"Blur","模糊强度"},{"Shadow","阴影强度"},{"Tint","颜色浓度"},
                {"Magnification","放大倍率"},{"Strength","边光强度"},{"Spread","边缘映色"},{"LightAngle","边光角度"},
                {"GlowWidth","泛光宽度"},{"RefractionStrength","边光折射"},{"DispersionStrength","色散强度"}};
            return names.ContainsKey(key)?names[key]:key;
        }
    }
}





