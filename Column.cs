using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converter
{
    class Column
    {
        public String Name;
        public int No;
        public Boolean Key = false;

        public Column() { }

        public Column(int number)
        {
            No = number;
        }

        public Column(String name, int number):this(number)
        {
            Name = name;
        }
    }
}
