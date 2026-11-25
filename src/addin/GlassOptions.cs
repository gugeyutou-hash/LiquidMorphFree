using System;
using System.Drawing;

namespace LiquidMorph.Addin
{
    /// <summary>Parameters for one glass render. All defaults produce a pleasant result.</summary>
    public sealed class GlassOptions
    {
        public int Material = 18;
        public int Red = 255, Green = 255, Blue = 255;
        public double Blur, Shadow, Magnification;
        public double PatternScale = 1, PatternDepth = 1, PatternAngle;
        public double Strength = 0.65;
        public double GlowWidth = 0.6;
        public double Spread = 0.5;
        public double LightAngle = 135;
        public int EdgeType = 1;
        public double Refraction = 0.55;
        public double RefractionStrength = 0.6;
        public double Dispersion = 0.35;
        public double DispersionStrength = 0.4;
        public double Tint = 0.1;
        public double RadiusFactor = 0.12;   // radius = min(width,height) * factor
        public double Padding = 40;          // points around the shape
        public int Density = 2;              // supersampling
        public bool ReplaceShape = true;

        public GlassOptions()
        {
            ApplyMaterial(18);
            Shadow = 12;
            Tint = 0;
            LightAngle = 0;
            EdgeType = 0;
        }

        private static double Scale(int position, double min, double max)
        {
            return min + (position-1)*(max-min)/99;
        }

        internal double ShaderPatternAngle()
        {
            if(Material==5) return PatternAngle/100;
            if(Material==6 || Material==16) return PatternAngle;
            return PatternAngle*Math.PI/180;
        }

        public void ApplyMaterial(int id)
        {
            Material=id;
            Blur=id==2 ? 8 : 0;
            Refraction=(id==1 || id==3 || id>=4 && id!=16) ? 1 : 0;
            Dispersion=id==3 || id==18 ? 1 : 0;
            int glow=10, width=60;
            switch(id)
            {
                case 0: glow=50;width=30;break;
                case 1: width=40;break;
                case 2: glow=15;width=20;break;
                case 3: glow=3;width=30;break;
                case 4: glow=15;break;
                case 5: glow=30;width=30;break;
                case 12: width=20;break;
                case 18: glow=20;width=30;break;
            }
            Strength=Scale(glow,0,10);
            GlowWidth=Scale(width,.1,24);
            Spread=id==0?Scale(30,1,30):id==1?Scale(60,1,30):
                id==2||id==4||id==5||id==18?Scale(20,1,30):1;
            RefractionStrength=id==18?Scale(75,0,1.78787875):id==1||id==4?Scale(50,0,1.78787875):
                id==3?Scale(30,0,1.78787875):id==5?Scale(60,0,1.78787875):1;
            DispersionStrength=id==3||id==18?Scale(40,0,3):1;
            PatternScale=id==4?.25:id==12?Scale(40,.25,4):id==5?4:1;
            PatternDepth=id==4?Scale(50,0,3):id==12?3:id==5?Scale(40,0,3):1;
            PatternAngle=id==5?49.49495:id==6||id==16?1:0;
        }
    }

    public sealed class GlassResult
    {
        public int Kind;
        public string ImagePath;
        public double Left;
        public double Top;
        public double Width;
        public double Height;
    }
}
