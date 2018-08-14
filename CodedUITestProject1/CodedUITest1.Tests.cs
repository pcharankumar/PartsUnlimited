using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace CodedUITestProject1
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

        [TestMethod]
        public void CodedUITestMethod1()
        {
            var PUBrowser = BrowserWindow.Launch("https://pupsi.azurewebsites.net/");
            if (PUBrowser.Exists)
            {
                TestContext.WriteLine("Parts Unlimited Browser exits");
            }
            else
            {
                TestContext.WriteLine("Parts Unlimited Browser does not exits");
            }

            PUBrowser.DrawHighlight();
            Playback.Wait(5000);
        }

        [TestMethod]
        public void CodedUITestMethod2()
        {
            var PUBrowser = BrowserWindow.Launch("https://pupsi.azurewebsites.net/");
            this.UIMap.SearchBrakes();

            Playback.Wait(5000);
        }

        [TestMethod]
        public void CodedUITestMethod3()
        {
            var PUBrowser = BrowserWindow.Launch("https://pupsi.azurewebsites.net/");
            Playback.Wait(5000);
        }

        [TestMethod]
        public void SearchParts()
        {
            var PUBrowser = BrowserWindow.Launch("https://pupsi.azurewebsites.net/");
            Playback.Wait(5000);
            this.UIMap.SearchParts();

        }

        [TestMethod]
        public void VerifyMenus()
        {
            var PUBrowser = BrowserWindow.Launch("https://pupsi.azurewebsites.net/");
            Playback.Wait(5000);
            this.UIMap.VerifyMenuOptions();


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
                if (this.map == null)
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}
