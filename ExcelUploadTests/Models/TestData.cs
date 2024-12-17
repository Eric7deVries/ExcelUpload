using Bogus;

namespace ExcelUploadTests.Models;

public class TestData : ICloneable
{
	public TestData()
	{
		People = [];
	}

	public class Person : ICloneable
	{
		public long ID { get; set; }
		public string Name { get; set; }
		public int Age { get; set; }

		public static Faker<Person> FakeData { get; } =
		new Faker<Person>()
			.RuleFor(p => p.ID, f => f.Random.Long(1, 1000))
				.RuleFor(p => p.Name, f => f.Name.FullName())
				.RuleFor(p => p.Age, f => f.Random.Number(8, 90));

		public override bool Equals(object? obj)
		{
			return obj is Person person &&
				   ID == person.ID &&
				   Name.Equals(person.Name) &&
				   Age == person.Age;
		}

		public object Clone()
		{
			return new Person
			{
				ID = this.ID,
				Age = this.Age,
				Name = this.Name
			};
		}
	}

	public List<Person> People { get; set; }
	public Person TestSubject1 { get; set; }
	public Person TestSubject2 { get; set; }
	public Person TestSubject3 { get; set; }
	public Person TestSubject4 { get; set; }

	public object Clone()
	{
		var subject1 = (Person)this.TestSubject1.Clone();
		var subject2 = (Person)this.TestSubject2.Clone();
		var subject3 = (Person)this.TestSubject3.Clone();
		var subject4 = (Person)this.TestSubject4.Clone();
		return new TestData
		{
			TestSubject1 = subject1,
			TestSubject2 = subject2,
			TestSubject3 = subject3,
			TestSubject4 = subject4,
			People = [subject1, subject2, subject3, subject4]
		};
	}
}
