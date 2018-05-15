using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace JsonToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename    = "movies.json"; // array of movie
            string projectPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string dataPath    = Path.Combine(projectPath, filename);
            string outputFile  = "movies.xlsx";
            string json        = File.ReadAllText(dataPath);
            
            Excel.Application xlApp = new Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook  = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            JArray items       = JArray.Parse(json);
            List<string> props = new List<string>();

            // get all properties
            foreach (JObject item in items)
            {
                foreach (JProperty prop in item.Properties())
                {
                    if (props.Contains(prop.Name))
                        continue;

                    props.Add(prop.Name);
                }
            }

            int rowStart = 1;

            // headers
            foreach (var p in props.Select((value, i) => new { i, value }))
            {
                xlWorkSheet.Cells[rowStart, p.i + 1] = p.value;
            }

            // data
            foreach (JObject item in items)
            {
                foreach (var prop in props.Select((value, i) => new { i, value }))
                {
                    if((dynamic)item[prop.value] != null)
                    {
                        xlWorkSheet.Cells[rowStart + 1, prop.i + 1] = (dynamic)item[prop.value];
                    }
                }

                rowStart++;
            }

            xlWorkBook.SaveAs(Path.Combine(projectPath, outputFile));
            //xlWorkBook.SaveAs(Path.Combine(projectPath, outputFile), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            // cleanup
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
