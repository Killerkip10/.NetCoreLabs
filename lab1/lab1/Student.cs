using System;
using System.Collections.Generic;
using System.Text;

namespace lab1
{
    class Student
    {
        public string name;
        public string group;
        public List<Semester> semesters = new List<Semester>();

        public Student(string name, string group)
        {
            this.name = name;
            this.group = group;
        }

        public void addSemester(Semester semester)
        {
            semesters.Add(semester);
        }
    }
}
