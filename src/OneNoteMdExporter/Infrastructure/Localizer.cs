﻿using Newtonsoft.Json.Linq;
using System;
using System.Globalization;
using System.IO;

namespace alxnbl.OneNoteMdExporter.Infrastructure
{
    public static class Localizer
    {
        public static string GetString(string code)
        {
            var lang = CultureInfo.CurrentCulture.TwoLetterISOLanguageName;//获取当前系统的文字
            var transFile = Path.Combine("Resources", $"trad.{lang}.json");
            var transFileEn = Path.Combine("Resources", $"trad.en.json");
            var transFileCn = Path.Combine("Resources", $"trad.zh.json");
            string localizedText = null;

            if (File.Exists(transFile))
            {
                var tradsFile = File.ReadAllText(transFile);
                JObject d = JObject.Parse(tradsFile);
                localizedText = d[code]?.ToString();
            }

            if(localizedText == null)
            {
                // Translation not found in current language

                var tradsFile = File.ReadAllText(transFileEn);//切换语言文件
                JObject d = JObject.Parse(tradsFile);

                if (d[code] != null)
                    localizedText = d[code].ToString();
                else
                    throw new InvalidOperationException($"Missing translation for code {code}");
            }

            return localizedText;
        }
    }
}
