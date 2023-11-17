using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.DataExtraction
{
    public class Csv
    {
        public static T[] Read<T>(string filePath, CultureInfo cultureInfo = null)
        {
            T[] records;

            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, cultureInfo ?? CultureInfo.InvariantCulture))
            {
                var recordsLazy = csv.GetRecords<T>();
                records = recordsLazy.ToArray();
            }

            return records;
        }

    }
}
