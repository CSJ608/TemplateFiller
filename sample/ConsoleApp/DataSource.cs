using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    public class DataSource
    {
        public string StartTime { get; set; } = string.Empty;
        public string EndTime { get; set; } = string.Empty;
        public string PrintTime { get; set; } = string.Empty;
        public Person[] Persons { get; set; } = [];
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Sex { get; set; } = string.Empty;
        public string WorkNo { get; set; } = string.Empty;
        public string Description {  get; set; } = string.Empty;
    }
}
