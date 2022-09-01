using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CIS.Controller
{
    class ClassImportCriteria
    {
        public static List<Model.Criteria> ReadCriteriaExcel(string filePath)
        {
            List<Model.Criteria> newListCriteria = new List<Model.Criteria>();

            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
            Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filePath);
            Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);//последнюю ячейку

            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;

            string[] massDataStudentFromExcel = new string[2];

            Model.CISEntities db = new Model.CISEntities();
            
            for (int i = 16; i <= 24; i++) //строки
            {
                for (int j = 2; j <= 11; j++) //столбцы
                {
                    massDataStudentFromExcel[j - 1] = ObjWorkSheet.Cells[i, j].Text.ToString();
                }

                Model.Criteria newCriteria = new Model.Criteria();

                newCriteria.Title = massDataStudentFromExcel[1];
                newCriteria.MaxValue = massDataStudentFromExcel[10];
                newCriteria.IdProModule = 1;

                newListCriteria.Add(newCriteria);
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из Excel
            GC.Collect(); // убрать за собой

            return newListCriteria;
        }
    }
}
