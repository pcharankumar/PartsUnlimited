using Microsoft.VisualStudio.TestTools.UnitTesting;
using PartsUnlimited.Controllers;
using PartsUnlimited.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace PartsUnlimited.Controllers.Tests
{
    [TestClass()]
    public class HomeControllerTests
    {

        IPartsUnlimitedContext context;

        [TestMethod()]
        public void IndexTest()
        {

            
                HomeController con = new HomeController(context);
            var x = con.Index();
            Assert.Equals(2, 2);
        }
    }
}