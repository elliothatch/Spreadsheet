﻿using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;


namespace SpreadsheetGUITest
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class CodedUITest1
    {
        public CodedUITest1()
        {
        }

        /// <summary>
        /// Select cells and make sure ui updates values, close wihtout saving
        /// </summary>
        [TestMethod]
        public void CodedUITestMethod1()
        {
            this.UIMap.Launch();
            this.UIMap.SelectCells1();
            this.UIMap.AssertSelectedCellD11();
            this.UIMap.SetA1Number1();
            this.UIMap.CloseWithoutSaving();
        }

        /// <summary>
        /// Test setting content, save, save as, close and save
        /// </summary>
        [TestMethod]
        public void CodedUITestMethod2()
        {
            this.UIMap.Launch();
            this.UIMap.SetA1Number1();
            this.UIMap.SetB1StringSelect1();
            this.UIMap.SetC1Formula1();
            this.UIMap.SetC4Formula1();
            this.UIMap.AssertC4Value();
            this.UIMap.SaveFirstTimeOverwriting();
            this.UIMap.SaveAsOverwriting();
            this.UIMap.SetA1Number1();
            this.UIMap.CloseAndCancel();
            this.UIMap.CloseAndSave();
        }

        /// <summary>
        /// Test opening and closing spreadsheets, help
        /// </summary>
        [TestMethod]
        public void CodedUITestMethod3()
        {
            this.UIMap.Launch();
            this.UIMap.OpenCloseHelp();
            this.UIMap.OpenSpreadsheet();
            this.UIMap.CloseNoPrompt1();
            this.UIMap.CloseOpenedWindowNoPrompt();
        }

        /// <summary>
        /// Test new spreadsheet, close with toolstrip
        /// </summary>
        [TestMethod]
        public void CodedUItestMethod4()
        {
            this.UIMap.Launch();
            this.UIMap.NewSpreadsheet();
            this.UIMap.CloseNewNoPromptToolstrip();
            this.UIMap.CloseNoPrompt1();
        }

        /// <summary>
        /// Test error handling
        /// </summary>
        [TestMethod]
        public void CodedUItestMethod5()
        {
            this.UIMap.Launch();
            this.UIMap.SetA1InvalidFormula1();
            this.UIMap.AssertErrorMessageDisplayed();
            this.UIMap.SetA1Number1();
            this.UIMap.SetA1CircularDependency();
            this.UIMap.AssertErrorMessageDisplayed();
            this.UIMap.CloseWithoutSaving();
        }

        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        ////Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        #endregion

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
        private TestContext testContextInstance;

        public UIMap UIMap
        {
            get
            {
                if ((this.map == null))
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}
