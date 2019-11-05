using System;
using System.Collections.Generic;
using System.Text;

namespace lab1
{
    class Semester
    {
        public int num;
        public List<Mark> marks = new List<Mark>();

        public Semester(int num)
        {
            this.num = num;
        }

        public void addMark(Mark mark)
        {
            marks.Add(mark);
        }
    }
}
