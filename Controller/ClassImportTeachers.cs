using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CIS.Controller
{
    class ClassImportTeachers
    {
        public static List<Model.Teachers> ReadExcelTeacher(string filePath)
        {
            //Метод считывания данных из Excel
            //Считанные данные помещаются в объявленный список
            List<Model.Teachers> newListTeachers = new List<Model.Teachers>();

            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
            Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filePath);
            Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;

            var CertainCell = ObjWorkSheet.Cells[1, 1].Text.ToString();
            if (CertainCell == "Преподаватели")
            {

                string[] massDataTeachersFromExcel = new string[6];

                Model.CISEntities db = new Model.CISEntities();

                for (int i = 3; i <= lastRow; i++)
                {
                    for (int j = 1; j <= lastColumn; j++)
                    {
                        massDataTeachersFromExcel[j - 1] = ObjWorkSheet.Cells[i, j].Text.ToString();
                    }

                    Model.Teachers newTeacher = new Model.Teachers();

                    newTeacher.FirstName = massDataTeachersFromExcel[0];
                    newTeacher.LastName = massDataTeachersFromExcel[1];
                    newTeacher.Patronymic = massDataTeachersFromExcel[2];
                    newTeacher.Email = massDataTeachersFromExcel[3];

                    string nameTeachstus = massDataTeachersFromExcel[4];
                    var status = db.StatusTeacher.Where(x => x.Title == nameTeachstus).FirstOrDefault();
                    newTeacher.IdStatusTeachers = status.IdStatusTeacher;

                    string nameTeachRole = massDataTeachersFromExcel[5];
                    var role = db.Role.Where(d => d.Title == nameTeachRole).FirstOrDefault();
                    newTeacher.IdRole = role.IdRole;

                    newListTeachers.Add(newTeacher);
                }
            }
            else
            {
                MessageBox.Show("Выбран неверный тип файла!\nВыберите тип файла.");
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            return newListTeachers;
        }
    }
}
