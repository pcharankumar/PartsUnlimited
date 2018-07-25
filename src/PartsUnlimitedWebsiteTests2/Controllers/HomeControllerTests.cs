using Microsoft.VisualStudio.TestTools.UnitTesting;
using PartsUnlimited.Controllers;
using PartsUnlimited.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NSubstitute;
using PartsUnlimited.Models;
using System.Web.Mvc;

namespace PartsUnlimited.Controllers.Tests
{
    [TestClass()]
    public class HomeControllerTests
    {
        Models.IPartsUnlimitedContext newContext;
        [TestMethod()]
        public void IndexTest()
        {
            var newContext = Substitute.For<IPartsUnlimitedContext>();
            HomeController hc = new HomeController(newContext);
            
            var result = hc.Recomendations() as ViewResult;
            //Assert.AreEqual(result, hc.Index());
        }
    }
}