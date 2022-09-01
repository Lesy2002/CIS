using CIS.Controller;
using Microsoft.Win32;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CIS
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<Model.Student> listStudents;
        public List<Model.Teachers> listTeachers;
        public List<Model.Criteria> listCriteria;

        public MainWindow()
        {
            InitializeComponent();
        }

        private string filePath;
        DateTime time = DateTime.Now;
        int count = 0;
        int countLine = 0;

        private void FileImportBtn_Click(object sender, RoutedEventArgs e)
        {
            //Метод открытия диалогового окна для выбора файла.
            if (StudentRadBtn.IsChecked == true || TeacherRadBtn.IsChecked == true || CriteriaRadBtn.IsChecked == true)
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.DefaultExt = "*.xls;*.xlsx";
                ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";

                if (ofd.ShowDialog() == true)
                {
                    filePath = ofd.FileName;
                    AddToDbBtn.IsEnabled = true;
                    messageText.Text += (count += 1) + "-" + time + "-" + "Файл выбран" + ":" + " " + ofd.FileName + "\n";
                }
                else
                {
                    AddToDbBtn.IsEnabled = false;
                }
            }
            else
            {
                messageText.Text += (count += 1) + "-" + time + "-Выберите тип файла!\n";
            }

            if (listStudents != null)
            {
                listStudents.Clear();
            }
            if (listTeachers != null)
            {
                listTeachers.Clear();
            }
        }

        private void AddToDbBtn_Click(object sender, RoutedEventArgs e)
        {
            Model.CISEntities db = new Model.CISEntities();

            if (StudentRadBtn.IsChecked == true)
            {
                listStudents = ClassImportStudent.ReadExcel(filePath);

                if (listStudents.Count != 0)
                {
                    int importSum = 0;
                    foreach (Model.Student student in listStudents)
                    {
                        var findStudent = db.Student.Where(x => x.Email == student.Email).FirstOrDefault();

                        if (findStudent == null)
                        {
                            importSum += 1;
                            db.Student.Add(student);
                            db.SaveChanges();

                        }
                        else
                        {
                            messageText.Text += (count += 1) + "-" + time + "-У студента " + student.FirstName + " " + student.LastName
                                + " " + student.Patronymic + " " + "не уникальный email. Данный студент не будет добавлен в БД!\n";
                        }
                    }

                    if (listTeachers != null)
                    {
                        countLine = Convert.ToInt32(listTeachers.Count);
                    }
                    if (listStudents != null)
                    {
                        countLine = Convert.ToInt32(listStudents.Count);
                    }
                    messageText.Text += (count += 1) + "-" + time + "-Количество считанных строк: " + countLine + "\n";
                    messageText.Text += (count += 1) + "-" + time + "-Число импортируемых строк " + importSum + "\n";
                    messageText.Text += (count += 1) + "-" + time + "-Данные успешно сохранены!\n";
                    AddToDbBtn.IsEnabled = false;
                }
                else
                {
                    messageText.Text += (count += 1) + "-" + time + "-Выберите файл для сохранения!\n";
                }
            }

            else if (TeacherRadBtn.IsChecked == true)
            {
                listTeachers = ClassImportTeachers.ReadExcelTeacher(filePath);

                if (listTeachers.Count != 0)
                {
                    int importSum = 0;
                    foreach (Model.Teachers teachers in listTeachers)
                    {
                        var findTeacher = db.Teachers.Where(x => x.Email == teachers.Email).FirstOrDefault();

                        if (findTeacher == null)
                        {
                            importSum += 1;
                            db.Teachers.Add(teachers);
                            db.SaveChanges();
                        }
                        else
                        {
                            messageText.Text += (count += 1) + "-" + time + "-У преподавателя " + teachers.FirstName + " " + teachers.LastName
                                + " " + teachers.Patronymic + " " + "не уникальный email. Данный преподаватель не будет добавлен в БД!\n";
                        }
                    }

                    if (listTeachers != null)
                    {
                        countLine = Convert.ToInt32(listTeachers.Count);
                    }
                    if (listStudents != null)
                    {
                        countLine = Convert.ToInt32(listStudents.Count);
                    }
                    messageText.Text += (count += 1) + "-" + time + "-Количество считанных строк: " + countLine + "\n";
                    messageText.Text += (count += 1) + "-" + time + "-Число импортируемых строк " + importSum + "\n";
                    messageText.Text += (count += 1) + "-" + time + "-Данные успешно сохранены!\n";
                    AddToDbBtn.IsEnabled = false;
                }
                else
                {
                    messageText.Text += (count += 1) + "-" + time + "-Выберите файл для сохранения!\n";
                }
            }
            if (listStudents != null)
            {
                listStudents.Clear();
            }
            if (listTeachers != null)
            {
                listTeachers.Clear();
            }
        }

        private void TeacherRadBtn_Checked(object sender, RoutedEventArgs e)
        {
            if (listStudents != null)
            {
                listStudents.Clear();
            }
            messageText.Text += (count += 1) + "-" + time + "-Выбран тип файла (преподаватели) \n";
        }

        private void StudentRadBtn_Checked(object sender, RoutedEventArgs e)
        {
            if (listTeachers != null)
            {
                listTeachers.Clear();
            }
            messageText.Text += (count += 1) + "-" + time + "-Выбран тип файла (студенты) \n";
        }

        private void CriteriaRadBtn_Checked(object sender, RoutedEventArgs e)
        {
            messageText.Text += (count += 1) + "-" + time + "-Выбран тип файла (студенты) \n";
        }
    }
}


