using ExcelGenerator.Helper;
using OfficeOpenXml;
using System;
using System.Collections;
using System.ComponentModel;
using System.Reflection;

namespace ExcelGenerator
{
    public static class ExcelHelper
    {
        public static byte[] WriteTsv<T>(T data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));

            using (ExcelPackage package = new ExcelPackage())
            {
                foreach (PropertyInfo propertyInfo in data.GetType().GetProperties())
                {
                    BuildServiceExcel(package, propertyInfo, propertyInfo.GetValue(data, null));
                }

                return package.GetAsByteArray();
            }
        }

        private static void BuildServiceExcel(ExcelPackage package, PropertyInfo propertyInfo, object p)
        {
            if (p != null)
            {
                var obj = (IEnumerable)(p);

                ExcelWorksheet ws = package.Workbook.Worksheets.Add(Helpers.GetDisplayName(propertyInfo));

                ws.View.ShowGridLines = true;

                int indexLine = 1;

                foreach (var exportLead in obj)
                {
                    BuildCells(ws, exportLead, ref indexLine);
                    indexLine++;
                }
            }
        }

        private static object GetValue(PropertyInfo propertyInfoLead, object obj)
        {
            if (propertyInfoLead.PropertyType == typeof(DateTime))
            {
                return Convert.ToDateTime(propertyInfoLead.GetValue(obj)).ToString("dd/MM/yyyy hh:mm");
            }
            else if (propertyInfoLead.PropertyType == typeof(DateTime?))
            {
                if (propertyInfoLead.GetValue(obj) != null)
                {
                    return Convert.ToDateTime(propertyInfoLead.GetValue(obj)).ToString("dd/MM/yyyy hh:mm");
                }
            }

            return propertyInfoLead.GetValue(obj);
        }

        private static void BuildCells(ExcelWorksheet ws, object obj, ref int indexLine)
        {
            int indexColumn = 1;

            if (indexLine == 1)
            {
                foreach (PropertyInfo propertyInfoLead in obj.GetType().GetProperties())
                {
                    bool hidden = Helpers.HiddenAttribute(propertyInfoLead);

                    if (!hidden)
                    {
                        ws.Cells[1, indexColumn].Value = Helpers.GetDisplayName(propertyInfoLead);
                        indexColumn++;
                    }
                }

                indexLine++;
            }

            indexColumn = 1;

            foreach (PropertyInfo propertyInfoLead in obj.GetType().GetProperties())
            {
                string name = Helpers.GetDisplayName(propertyInfoLead);
                bool hidden = Helpers.HiddenAttribute(propertyInfoLead);

                if (!hidden)
                {
                    ws.Cells[indexLine, indexColumn].Value = GetValue(propertyInfoLead, obj);

                    indexColumn++;
                }
            }
        }        
    }
}
