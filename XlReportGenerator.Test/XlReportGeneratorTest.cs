using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator.Test
{
	[TestClass]
    public class XlReportGeneratorTest
    {
		[TestMethod]
        public void TestGeneratedRandomFileNameWithSimpleClass1AsData()
        {
            SimpleClass1 data = new SimpleClass1()
            {
                Field1 = "Field 1 Line 3",
                Field2 = "Field 2 Line 3"
            };

            XlReportGenerator.Generate(data, "D:\\Test", "Report 1234");
        }

        [TestMethod]
        public void TestGeneratedRandomFileNameWithSimpleClass2AsData()
        {
            SimpleClass2 data = new SimpleClass2()
            {
                Name = "Tigabelas",
                Age = 20,
                BOD = new DateTime(1994, 04, 20)
            };

            XlReportGenerator.Generate(data, "D:\\Test", "Report 1235");
        }

        [TestMethod]
        public void TestGeneratedRandomFileNameWithListSimpleClass2AsData()
        {
            List<SimpleClass1> lstSimpleClass1 = new List<SimpleClass1>();


            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 1",
                Field2 = "Field 2 Line 1"
            });

            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 2",
                Field2 = "Field 2 Line 2"
            });

            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 3",
                Field2 = "Field 2 Line 3"
            });


            XlReportGenerator.Generate(lstSimpleClass1, "D:\\Test", "Report 1235");
        }

        [TestMethod]
        public void TestGeneratedRandomFileNameWithComplexClass1AsData()
        {
            ComplexClass1 data = new ComplexClass1()
            {
                SC1 = new SimpleClass1()
                {
                    Field1 = "Hello",
                    Field2 = "World"
                },
                SC2 = new SimpleClass2()
                {
                    Name = "Tigabelas",
                    Age = 20,
                    BOD = new DateTime(1994, 04, 20)
                },
                SC3 = "Hello"
            };

            XlReportGenerator.Generate(data, "D:\\Test", "Report 1235");
        }

        [TestMethod]
        public void TestGeneratedRandomFileNameWithComplexClass2AsData()
        {

            List<SimpleClass1> lstSimpleClass1 = new List<SimpleClass1>();


            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 1",
                Field2 = "Field 2 Line 1"
            });

            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 2",
                Field2 = "Field 2 Line 2"
            });

            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 3",
                Field2 = "Field 2 Line 3"
            });

            ComplexClass2 data = new ComplexClass2()
            {
                SC0 = "Hello 0",
                SC1 = lstSimpleClass1,
                SC2 = "Hello 3"
            };

            XlReportGenerator.Generate(data, "D:\\Test", "Report 1235");
        }

        [TestMethod]
        public void TestGeneratedRandomFileNameWithListComplexClass2AsData()
        {

            List<ComplexClass2> lstComplexClass2 = new List<ComplexClass2>();
            List<SimpleClass1> lstSimpleClass1 = new List<SimpleClass1>();


            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 1",
                Field2 = "Field 2 Line 1"
            });

            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 2",
                Field2 = "Field 2 Line 2"
            });

            lstSimpleClass1.Add(new SimpleClass1
            {
                Field1 = "Field 1 Line 3",
                Field2 = "Field 2 Line 3"
            });

            lstComplexClass2.Add(new ComplexClass2()
            {
                SC0 = "Hello 0",
                SC1 = lstSimpleClass1,
                SC2 = "Hello 30"
            });

            lstComplexClass2.Add(new ComplexClass2()
            {
                SC0 = "Hello 1",
                SC1 = lstSimpleClass1,
                SC2 = "Hello 31"
            });

            lstComplexClass2.Add(new ComplexClass2()
            {
                SC0 = "Hello 2",
                SC1 = lstSimpleClass1,
                SC2 = "Hello 32"
            });

            XlReportGenerator.Generate(lstComplexClass2, "D:\\Test", "Report 1235");
        }
    }
}
