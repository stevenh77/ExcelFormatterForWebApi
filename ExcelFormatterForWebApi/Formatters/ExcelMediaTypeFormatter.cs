using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Reflection;
using System.Web;

namespace ExcelFormatterForWebApi.Formatters
{
    public class ExcelMediaTypeFormatter : BufferedMediaTypeFormatter
    {
        private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //private readonly Func<T, ExcelRow> builder;
        //Func<T, ExcelRow> value
        public ExcelMediaTypeFormatter()
        {
            //builder = value;
            SupportedMediaTypes.Add(new MediaTypeHeaderValue(ContentType));
            this.MediaTypeMappings.Add(new UriPathExtensionMapping("xlsx", ContentType));
        }

        public override bool CanWriteType(Type type)
        {
            return true; // type == typeof(IQueryable<T>) || type == typeof(T);
        }

        public override bool CanReadType(Type type)
        {
            return false;
        }
        public override void WriteToStream(Type type, object value, Stream writeStream, HttpContent content)
        {
            using (Stream ms = new MemoryStream())
            {
                using (var book = new ClosedXML.Excel.XLWorkbook())
                {
                    var sheet = book.Worksheets.Add("sample");
                    if (value is IEnumerable)
                    {
                        var rowIndex = 0;
                        foreach (var item in value as IEnumerable)
                        {
                            var columnIndex = 0;
                            foreach (PropertyInfo prop in item.GetType().GetProperties())
                            {
                                sheet.Cell(rowIndex +2, columnIndex+2).Value = prop.GetValue(item, null);
                                columnIndex++;
                            }                            
                            rowIndex++;
                        }
                    }
                    else
                    {
                        // output an individual row
                    }

                    sheet.Columns().AdjustToContents();

                    book.SaveAs(ms);

                    byte[] buffer = new byte[ms.Length];

                    ms.Position = 0;
                    ms.Read(buffer, 0, buffer.Length);

                    writeStream.Write(buffer, 0, buffer.Length);
                }
            }
        }
    }
}