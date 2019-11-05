using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace lab1
{
    class MarksContainer
    {
        public List<Student> students = new List<Student>();
        public List<string> title;

        public MarksContainer(string[] subjects)
        {
            title = new List<string>(subjects);
        }

        public void addStudent(Student student)
        {
            students.Add(student);
        }

        public List<double> getGradePointAverage(int index)
        {
            List<double> average = new List<double>();
            Student student = students[index];
            int markSum = 0;

            for (int i = 0; i < student.semesters.Count; i++, markSum = 0)
            {
                for(int k = 0; k < student.semesters[i].marks.Count; k++)
                {
                    markSum += student.semesters[i].marks[k].mark;
                }

                average.Add(markSum / student.semesters[i].marks.Count);
            }

            return average;
        }

        public List<double> getGradePointAverageSubject()
        {
            List<double> average = new List<double>();
            double markSum = 0;
            int semestersCount = 0;

            for(int i = 0; i < title.Count - 3; i++, markSum = 0)
            {
                for(int k = 0; k < students.Count; k++)
                {
                    semestersCount += students[k].semesters.Count;

                    for (int j = 0; j < students[k].semesters.Count; j++)
                    {
                        markSum += students[k].semesters[j].marks[i].mark;
                    }
                }

                average.Add(markSum / semestersCount);
                semestersCount = 0;
            }

            return average;
        }

        public string getJSON()
        {
            return JsonConvert.SerializeObject(this);
        }

        public ExcelPackage getXLSX(string name)
        {
            ExcelPackage excel = new ExcelPackage();
            List<double> average;
            List<double> averageSubjects = getGradePointAverageSubject();
            int line = 2;

            excel.Workbook.Worksheets.Add(name);

            var worksheet = excel.Workbook.Worksheets[name];

            for (int i = 0; i < title.Count; i++)
            {
                worksheet.Cells[
                    Char.ConvertFromUtf32((i + 1) + 64) + "1:" + Char.ConvertFromUtf32((i + 2) + 64) + "1"
                ].LoadFromText(title[i]);
            }

            worksheet.Cells[
                Char.ConvertFromUtf32((title.Count + 1) + 64) + "1:" + Char.ConvertFromUtf32((title.Count + 2) + 64) + "1"
            ].LoadFromText("Average");

            for (int i = 0; i < students.Count; i++)
            {
                worksheet.Cells[
                   "A" + line + ":" + "B" + line
                ].LoadFromText(students[i].name);
                worksheet.Cells[
                   "B" + line + ":" + "C" + line
                ].LoadFromText(students[i].group);

                average = getGradePointAverage(i);

                for (int k = 0, j = 0; k < students[i].semesters.Count; k++)
                {
                    worksheet.Cells[
                        "C" + line + ":" + "D" + line
                    ].LoadFromText(students[i].semesters[k].num.ToString());
                    

                    for(j = 0; j < students[i].semesters[k].marks.Count; j++)
                    {
                        worksheet.Cells[
                            Char.ConvertFromUtf32(68 + j) + line + ":" + Char.ConvertFromUtf32(68 + j) + line
                        ].LoadFromText(students[i].semesters[k].marks[j].mark.ToString());
                    }

                    worksheet.Cells[
                        Char.ConvertFromUtf32(68 + j) + line + ":" + Char.ConvertFromUtf32(69 + j) + line
                    ].LoadFromText(average[k].ToString());

                    line++;
                }
                line++;
            }

            for (int i = 0; i < averageSubjects.Count; i++)
            {
                worksheet.Cells[
                    Char.ConvertFromUtf32(68 + i) + line + ":" + Char.ConvertFromUtf32(68 + i) + line
                ].LoadFromText(averageSubjects[i].ToString());
            }

            worksheet.Cells[
                   Char.ConvertFromUtf32(68 + averageSubjects.Count) + line + ":" + Char.ConvertFromUtf32(68 + averageSubjects.Count) + line
               ].LoadFromText("");


            return excel;
        }
    }
}
