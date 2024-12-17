using Bogus;
using ExcelUpload.Services;
using ExcelUploadTests.Models;

namespace ExcelUploadTests.Factories;

public static class TestFactory
{
	public static TestData CreateTestData()
	{
		TestData data = new TestData();
		var testSubject1 = TestData.Person.FakeData.Generate();

		data.People.Add(testSubject1);
		data.TestSubject1 = testSubject1;

		var testSubject2 = TestData.Person.FakeData.Generate();

		data.People.Add(testSubject2);
		data.TestSubject2 = testSubject2;

		var testSubject3 = TestData.Person.FakeData.Generate();

		data.People.Add(testSubject3);
		data.TestSubject3 = testSubject3;

		var testSubject4 = TestData.Person.FakeData.Generate();

		data.People.Add(testSubject4);
		data.TestSubject4 = testSubject4;

		return data;
	}
}
