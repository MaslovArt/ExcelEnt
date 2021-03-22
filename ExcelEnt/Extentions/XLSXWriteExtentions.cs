using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelEnt.Write
{
    public static class XLSXWriteExtentions
    {
        public static void SaveTo(this XSSFWorkbook workbook, string filePath)
        {
            using (var file = new FileStream(filePath, FileMode.CreateNew))
            {
                workbook.Write(file, true);
            }
        }
    }
}
