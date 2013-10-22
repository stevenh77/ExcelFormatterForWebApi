using System.Net.Http.Formatting;

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