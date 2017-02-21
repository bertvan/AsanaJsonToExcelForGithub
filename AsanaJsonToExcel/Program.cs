using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsanaJsonToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var filePath = @"c:\temp\edana.json";

            var jsonObject = JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(filePath));

            var results = new List<dynamic>();

            foreach (var jsonTask in jsonObject["data"])
            {
                var excelRow = new
                {
                    Name = (string)jsonTask["name"].Value,
                    Completed = jsonTask["completed"].Value,
                    Estimate = jsonTask["custom_fields"]?[0]["number_value"].Value,
                    DevState = ExtractDevState(jsonTask),
                };

                results.Add(excelRow);
            }

            using (ExcelPackage pck = new ExcelPackage())
            {
                //Create the worksheet
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Demo");

                ws.Cells[1, 1].Value = "Name";
                ws.Cells[1, 2].Value = "Completed";
                ws.Cells[1, 3].Value = "Estimate";
                ws.Cells[1, 4].Value = "DevState";

                for (int i = 0; i < results.Count; i++)
                {
                    var item = results[i];
                    ws.Cells[i + 2, 1].Value = item.Name;
                    ws.Cells[i + 2, 2].Value = item.Completed;
                    ws.Cells[i + 2, 3].Value = item.Estimate;
                    ws.Cells[i + 2, 4].Value = item.DevState;
                }

                pck.SaveAs(new FileInfo(@"c:\temp\asana-export.xlsx"));
            }

            var bla = "";
        }

        private static string ExtractDevState(dynamic jsonTask)
        {
            var enumValue = jsonTask["custom_fields"]?[2]?["enum_value"];

            if (enumValue == null)
            {
                return "";
            }

            return (string)enumValue?["name"].Value;
        }
    }
}
