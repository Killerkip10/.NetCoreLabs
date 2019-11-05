using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace lab1
{
    class Reader
    {
        MarksContainer marksContainer;

        public void read(string path)
        {
            using (StreamReader sr = new StreamReader(path, Encoding.Default))
            {
                string line;
                string[] words;
                string[] marks;
                string[] title;
                Student student = new Student("", "");
                Semester semester;
                Mark mark;

                title = sr.ReadLine().Split(";", StringSplitOptions.RemoveEmptyEntries);

                marksContainer = new MarksContainer(title);

                while ((line = sr.ReadLine()) != null)
                {
                    words = line.Split(";");

                    if (words[0] != "")
                    {
                        student = new Student(words[0], words[1]);
                        marksContainer.addStudent(student);
                    }

                    semester = new Semester(Convert.ToInt32(words[2]));

                    for(int i = 3; i < words.Length; i++)
                    {
                        mark = new Mark(Convert.ToInt32(words[i]), title[i]);
                        semester.addMark(mark);
                    }

                    student.addSemester(semester);
                }
            }
        }

        public void writeXLSX(string fileName)
        {
            ExcelPackage excel = marksContainer.getXLSX(fileName);

            excel.SaveAs(new FileInfo(Directory.GetCurrentDirectory() + "\\" + fileName));
        }

        public void writeJSON(string fileName)
        {
            string json = marksContainer.getJSON();

            using (StreamWriter newTask = new StreamWriter(fileName, false))
            {
                newTask.WriteLine(json);
            }
        }

        static public void log(string message)
        {
            using (StreamWriter logs = new StreamWriter("logs.txt", true))
            {
                logs.WriteLine(message);
            }
        }
    }
}
