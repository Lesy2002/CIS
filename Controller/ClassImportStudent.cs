using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace CIS.Controller
{
    class ClassImportStudent
    {
        public static List<Model.Student> ReadExcel(string filePath)
        {
            //Метод считывания данных из Excel
            //Считанные данные помещаются в объявленный список
            List<Model.Student> newListStudent = new List<Model.Student>();

            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
            Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filePath);
            Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;

            var CertainCell = ObjWorkSheet.Cells[1, 1].Text.ToString();
            if (CertainCell == "Студенты")
            {
                string[] massDataStudentFromExcel = new string[8];

                Model.CISEntities db = new Model.CISEntities();
                for (int i = 3; i <= lastRow; i++) //строка
                {
                    for (int j = 1; j <= lastColumn; j++) //столбец
                    {
                        massDataStudentFromExcel[j - 1] = ObjWorkSheet.Cells[i, j].Text.ToString();
                    }
                    Model.Student newStudent = new Model.Student();
                    
                    newStudent.FirstName = massDataStudentFromExcel[0];
                    newStudent.LastName = massDataStudentFromExcel[1];
                    newStudent.Patronymic = massDataStudentFromExcel[2];
                    newStudent.Email = massDataStudentFromExcel[3];
                    newStudent.Telephone = massDataStudentFromExcel[4];
                    
                    string idGr = massDataStudentFromExcel[5];
                    var group = db.Group.Where(x => x.Title == idGr).FirstOrDefault();
                    newStudent.IdGroup = group.IdGroup;

                    string nameSpec = massDataStudentFromExcel[6];
                    var spec = db.Speciality.Where(x => x.Title == nameSpec).FirstOrDefault();
                    newStudent.IdSpeciality = spec.IdSpeciaity;

                    string nameStstus = massDataStudentFromExcel[7];
                    var status = db.StatusStudent.Where(x => x.Title == nameStstus).FirstOrDefault();
                    newStudent.IdStatusStudent = status.IdStatusStudent;

                    newListStudent.Add(newStudent);
                }
            }
            else
            {
                MessageBox.Show("Выбран неверный тип файла!\nВыберите тип файла.");
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            return newListStudent;
        }
    }
}
