namespace ConsoleApp
{
    public class DataSource
    {
        public DataSource(Stream codeStream)
        {
            Code = codeStream;
        }

        public string StartTime { get; set; } = string.Empty;
        public string EndTime { get; set; } = string.Empty;
        public string PrintTime { get; set; } = string.Empty;
        public Person[] Persons { get; set; } = [];
        public int[] Counts { get; set; } = [];
        public Stream Code { get; set; }
    }

    public class Person
    {
        public PersonInfo Info { get; set; } = new PersonInfo();
        public string WorkNo { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    public class PersonInfo
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Sex { get; set; } = string.Empty;
    }
}
