using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Formatting;
using System.Web;

namespace ExcelFormatterForWebApi.Formatters
{
    public class FormatterConfig
    {
        public static void RegisterFormatters(MediaTypeFormatterCollection formatters)
        {
            formatters.Insert(0, new ExcelMediaTypeFormatter());
        }
    }
}