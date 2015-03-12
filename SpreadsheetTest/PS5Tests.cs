using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetUtilities;
using SS;

namespace SpreadsheetTest
{
    /// <summary>
    /// Summary description for PS5Tests
    /// </summary>
    [TestClass]
    public class PS5Tests
    {
        public PS5Tests()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        /// <summary>
        /// test isValid on setContents
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void Test1()
        {
            Spreadsheet s = new Spreadsheet(a => false, a => a, "default");
            s.GetCellContents("A1");
        }

        /// <summary>
        /// test isValid on getContents
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void Test2()
        {
            Spreadsheet s = new Spreadsheet(a => false, a => a, "default");
            s.SetContentsOfCell("A0", "hello");
        }

        /// <summary>
        /// test isValid on getValue
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void Test3()
        {
            Spreadsheet s = new Spreadsheet(a => false, a => a, "default");
            s.GetCellValue("Z9");
        }

        /// <summary>
        /// Test normalize(isValid())
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidNameException))]
        public void Test4()
        {
            Spreadsheet s = new Spreadsheet(a => a[0] == 'b', a => a.ToUpper(), "default");
            s.SetContentsOfCell("b1", "=a1");
        }

        /// <summary>
        /// Test normalize
        /// </summary>
        [TestMethod]
        public void Test5()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a.ToUpper(), "default");
            s.SetContentsOfCell("a1", "hello");
            Assert.AreEqual("hello", s.GetCellContents("a1"));
            Assert.AreEqual("hello", s.GetCellContents("A1"));
            s.SetContentsOfCell("bc1", "2.3");
            Assert.AreEqual(2.3, s.GetCellContents("bc1"));
            Assert.AreEqual(2.3, s.GetCellContents("BC1"));
            s.SetContentsOfCell("z0", "=b1 + BC1");
            Formula f1 = new Formula("B1 + BC1");
            Assert.AreEqual(f1, s.GetCellContents("z0"));
            Assert.AreEqual(f1, s.GetCellContents("Z0"));
        }

        /// <summary>
        /// Test GetCellValue
        /// </summary>
        [TestMethod]
        public void Test6()
        {
            Spreadsheet s = new Spreadsheet();
            s.SetContentsOfCell("a1", "hello");
            s.SetContentsOfCell("b1", "5.2");
            s.SetContentsOfCell("b2", "2.3");
            s.SetContentsOfCell("b3", "0.0");
            s.SetContentsOfCell("z0", "=9");
            s.SetContentsOfCell("z1", "=9*0.5");
            s.SetContentsOfCell("z2", "=b1 + b2");
            s.SetContentsOfCell("z3", "=3.1/0");
            s.SetContentsOfCell("z4", "=3.1/b3");
            s.SetContentsOfCell("z5", "=3.1/c1");
            s.SetContentsOfCell("z6", "=z1 + 2*z2");
            s.SetContentsOfCell("z7", "=z6 * z2 - z1/2");
            s.SetContentsOfCell("z8", "=z1 * z2 + z6 - z5");
            s.SetContentsOfCell("z9", "=z1 * z2 + z6 - z4");
            Assert.AreEqual("hello", (string)s.GetCellValue("a1"));
            Assert.AreEqual(5.2, (double)s.GetCellValue("b1"));
            Assert.AreEqual(2.3, (double)s.GetCellValue("b2"));
            Assert.AreEqual(0.0, (double)s.GetCellValue("b3"));
            Assert.AreEqual(9.0, (double)s.GetCellValue("z0"));
            Assert.AreEqual(9.0*0.5, (double)s.GetCellValue("z1"));
            Assert.AreEqual(5.2+2.3, (double)s.GetCellValue("z2"));
            Assert.IsInstanceOfType(s.GetCellValue("z3"), typeof(FormulaError));
            Assert.IsInstanceOfType(s.GetCellValue("z4"), typeof(FormulaError));
            Assert.IsInstanceOfType(s.GetCellValue("z5"), typeof(FormulaError));
            Assert.AreEqual((9.0 * 0.5) + 2.0 * (5.2 + 2.3), (double)s.GetCellValue("z6"));
            Assert.AreEqual(((9.0 * 0.5) + 2.0 * (5.2 + 2.3)) * (5.2 + 2.3) - (9.0 * 0.5) / 2.0, (double)s.GetCellValue("z7"));
            Assert.IsInstanceOfType(s.GetCellValue("z8"), typeof(FormulaError));
            Assert.IsInstanceOfType(s.GetCellValue("z9"), typeof(FormulaError));

            //make sure one large formula correctly evaluates the value of several unevaluated formulas
            //(internally each stores a value which is calculated on GetCellValue, so the above tests don't thoroughly test the recursive evaluation)

            s.SetContentsOfCell("s1", "=5.2");
            s.SetContentsOfCell("s2", "=2.3");
            s.SetContentsOfCell("t1", "=9*0.5");
            s.SetContentsOfCell("t2", "=s1 + s2");
            s.SetContentsOfCell("t6", "=t1 + 2*t2");
            s.SetContentsOfCell("t7", "=t6 * t2 - t1/2");
            Assert.AreEqual(((9.0 * 0.5) + 2.0 * (5.2 + 2.3)) * (5.2 + 2.3) - (9.0 * 0.5) / 2.0, (double)s.GetCellValue("t7"));
        }

        /// <summary>
        /// Read nonexistant file
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void Test7()
        {
            Spreadsheet s = new Spreadsheet("test8.xml", a => true, a => a, "default");
        }

        /// <summary>
        /// Read wrong version number
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void Test9()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a, "default");
            s.SetContentsOfCell("A1", "hello");
            s.Save("test9.xml");

            Spreadsheet s2 = new Spreadsheet("test9.xml", a => true, a => a, "1.1");
        }

        /// <summary>
        /// Save invalid file name
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void Test10()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a, "default");
            s.SetContentsOfCell("A1", "hello");
            s.Save("test10?.xml");
        }

        /// <summary>
        /// Read invalid cell name
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void Test11()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a, "default");
            s.SetContentsOfCell("A1", "hello");
            s.SetContentsOfCell("B1", "2.2");
            s.Save("test11.xml");

            Spreadsheet s2 = new Spreadsheet("text11.xml", a => a[0] == 'A', a => a, "default");
        }

        /// <summary>
        /// Read invalid variable name
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void Test12()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a, "default");
            s.SetContentsOfCell("A1", "=1.1 + B1 + B2 + B12");
            s.SetContentsOfCell("B1", "2.2");
            s.SetContentsOfCell("B2", "3.2");
            s.Save("test12.xml");

            Spreadsheet s2 = new Spreadsheet("text12.xml", a => a.Length == 2, a => a, "default");
        }

        /// <summary>
        /// Make sure both normalizers work
        /// </summary>
        [TestMethod]
        public void Test13()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a.ToUpper(), "default");
            s.SetContentsOfCell("a1", "hello");
            s.SetContentsOfCell("b1", "3.14");
            s.SetContentsOfCell("c1", "=a1*2");
            s.Save("test13.xml");

            Spreadsheet s2 = new Spreadsheet("test13.xml", a => true, a => a + "0", "default");
            Assert.AreEqual("hello", (string)s2.GetCellContents("A1"));
            Assert.AreEqual(3.14, (double)s2.GetCellContents("B1"));
            Assert.AreEqual(new Formula("A10*2"), (Formula)s2.GetCellContents("C1"));
        }

        /// <summary>
        /// General read/write test
        /// </summary>
        [TestMethod]
        public void Test14()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a.ToUpper(), "default");
            s.SetContentsOfCell("a1", "hello");
            s.SetContentsOfCell("b1", "5.2");
            s.SetContentsOfCell("b2", "2.3");
            s.SetContentsOfCell("b3", "0.0");
            s.SetContentsOfCell("z0", "=9");
            s.SetContentsOfCell("z1", "=9*0.5");
            s.SetContentsOfCell("z2", "=b1 + b2");
            s.SetContentsOfCell("z3", "=3.1/0");
            s.SetContentsOfCell("z4", "=3.1/b3");
            s.SetContentsOfCell("z5", "=3.1/c1");
            s.Save("test14.xml");

            Spreadsheet s2 = new Spreadsheet("test14.xml", a => true, a => a, "default");
            Assert.IsTrue(new HashSet<string>(s.GetNamesOfAllNonemptyCells()).SetEquals(new HashSet<string>(s2.GetNamesOfAllNonemptyCells())));
            Assert.AreEqual(s.GetCellContents("a1"), s2.GetCellContents("A1"));
            Assert.AreEqual(s.GetCellContents("b1"), s2.GetCellContents("B1"));
            Assert.AreEqual(s.GetCellContents("b2"), s2.GetCellContents("B2"));
            Assert.AreEqual(s.GetCellContents("b3"), s2.GetCellContents("B3"));
            Assert.AreEqual(s.GetCellContents("z0"), s2.GetCellContents("Z0"));
            Assert.AreEqual(s.GetCellContents("z1"), s2.GetCellContents("Z1"));
            Assert.AreEqual(s.GetCellContents("z2"), s2.GetCellContents("Z2"));
            Assert.AreEqual(s.GetCellContents("z3"), s2.GetCellContents("Z3"));
            Assert.AreEqual(s.GetCellContents("z4"), s2.GetCellContents("Z4"));
            Assert.AreEqual(s.GetCellContents("z5"), s2.GetCellContents("Z5"));
        }

        /// <summary>
        /// Check saved version
        /// </summary>
        [TestMethod]
        public void Test15()
        {
            Spreadsheet s = new Spreadsheet(a => true, a => a, @"A?00b\/");
            s.Save("test15.xml");

            Assert.AreEqual(@"A?00b\/", s.GetSavedVersion("test15.xml"));
        }

        /// <summary>
        /// Test Changed property
        /// </summary>
        [TestMethod]
        public void Test16()
        {
            Spreadsheet s = new Spreadsheet();
            Assert.IsFalse(s.Changed);
            s.GetCellContents("A1");
            s.GetCellValue("B1");
            Assert.IsFalse(s.Changed);

            s.SetContentsOfCell("A1", "hello");
            Assert.IsTrue(s.Changed);

            s.Save("test16.xml");
            Assert.IsFalse(s.Changed);
            try
            {
                s.SetContentsOfCell("B1", "=B1");
            }
            catch(CircularException)
            {

            }
            Assert.IsFalse(s.Changed);

            Assert.AreEqual(@"A?00b\/", s.GetSavedVersion("test15.xml"));
        }

        /// <summary>
        /// Reads an XML file that has a circular dependency
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void Test17()
        {
            string[] lines = {
                           @"<?xml version=""default"" encoding=""utf-8""?>",
                           @"<spreadsheet version=""default"">",
                           @"<cell>",
                           @"<name>A1</name>",
                           @"<contents>=B1</contents>",
                           @"</cell>",
                           @"<cell>",
                           @"<name>B1</name>",
                           @"<contents>=C1</contents>",
                           @"</cell>",
                           @"<cell>",
                           @"<name>C1</name>",
                           @"<contents>=A1</contents>",
                           @"</cell>",
                           @"</spreadsheet>"
                       };
            System.IO.File.WriteAllLines(@"test17.xml", lines);
            Spreadsheet s = new Spreadsheet("test17.xml", a => true, a => a, "default");
        }

        /// <summary>
        /// Check saved version of nonexistant file
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void Test18()
        {
            Spreadsheet s = new Spreadsheet();
            s.GetSavedVersion("test18.xml");
        }

        /// <summary>
        /// GetCellValue: return a formula error when an equation is invalid
        /// </summary>
        [TestMethod]
        public void Test19()
        {
            Spreadsheet s = new Spreadsheet();
            s.SetContentsOfCell("A1", "hello");
            s.SetContentsOfCell("B1", "=A1 + 2");
            s.SetContentsOfCell("C1", "=B1 - 3");
            Assert.IsInstanceOfType(s.GetCellValue("C1"), typeof(FormulaError));
        }
    }
}
