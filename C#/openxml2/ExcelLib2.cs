using System.Runtime.InteropServices;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace openxml2
{
    public class ExcelLib2 : IExcelLib2
    {
        private static ExcelLib2 instance = new ExcelLib2();
        [return: MarshalAs(UnmanagedType.IUnknown)]
        public static IExcelLib2 ACreateInstance()
        {
            if (instance == null)
            {
                instance = new ExcelLib2();
            }
            return instance;
        }

        public void Test(out uint result)
        { 
            SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open("G:\\Vorlagendatei.xlsm", true); // Pfad ändern!
            WorkbookStylesPart stylesPart = spreadSheet.WorkbookPart.WorkbookStylesPart;
            Stylesheet stylesheet = stylesPart.Stylesheet;
            spreadSheet.Close();
            result = 23;
        }

        [DllExport("OpenExcelFile", CallingConvention = CallingConvention.StdCall)]
        public static void OpenExcelFile(
        [MarshalAs(UnmanagedType.Interface)] out IExcelLib2 instance)
        {
            if (ExcelLib2.instance == null)
            {
                ExcelLib2.instance = new ExcelLib2();
            }
            instance = ExcelLib2.instance;
        }
    }
}
