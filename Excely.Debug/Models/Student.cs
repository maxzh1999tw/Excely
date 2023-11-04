namespace Excely.Debug.Models
{
    internal class Student
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
        public DateTime? Birthday { get; set; }

        public Student() { }

        public Student(int id, string name, DateTime? birthday)
        {
            Id = id;
            Name = name;
            Birthday = birthday;
        }
    }
}
