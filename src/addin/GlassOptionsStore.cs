using System;
using System.IO;
using Microsoft.Win32;

namespace LiquidMorph.Addin
{
    /// <summary>Loads/saves glass options from HKCU so the ribbon is stateless.</summary>
    public static class GlassOptionsStore
    {
        private const string KeyPath = @"Software\LiquidMorph\Free";

        public static GlassOptions Load()
        {
            GlassOptions o = new GlassOptions();
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (key == null) return o;
                    string saved = key.GetValue("Options") as string;
                    if (!string.IsNullOrEmpty(saved)) return Json.Deserialize<GlassOptions>(saved);
                    o.Strength = ReadDouble(key, "Strength", o.Strength);
                    o.GlowWidth = ReadDouble(key, "GlowWidth", o.GlowWidth);
                    o.Spread = ReadDouble(key, "Spread", o.Spread);
                    o.LightAngle = ReadDouble(key, "LightAngle", o.LightAngle);
                    o.EdgeType = ReadInt(key, "EdgeType", o.EdgeType);
                    o.RefractionStrength = ReadDouble(key, "RefractionStrength", o.RefractionStrength);
                    o.DispersionStrength = ReadDouble(key, "DispersionStrength", o.DispersionStrength);
                    o.Tint = ReadDouble(key, "Tint", o.Tint);
                    o.Padding = ReadDouble(key, "Padding", o.Padding);
                    o.Density = ReadInt(key, "Density", o.Density);
                }
            }
            catch { }
            return o;
        }

        public static void Save(GlassOptions o)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(KeyPath))
                key.SetValue("Options", Json.Serialize(o));
        }

        internal static readonly System.Web.Script.Serialization.JavaScriptSerializer Json =
            new System.Web.Script.Serialization.JavaScriptSerializer();

        private static double ReadDouble(RegistryKey key, string name, double fallback)
        {
            object value = key.GetValue(name);
            if (value == null) return fallback;
            double parsed;
            return double.TryParse(Convert.ToString(value), out parsed) ? parsed : fallback;
        }

        private static int ReadInt(RegistryKey key, string name, int fallback)
        {
            object value = key.GetValue(name);
            if (value == null) return fallback;
            int parsed;
            return int.TryParse(Convert.ToString(value), out parsed) ? parsed : fallback;
        }
    }
}
