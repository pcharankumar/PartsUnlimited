using Microsoft.VisualStudio.TestTools.UnitTesting;
using NSubstitute;
using PartsUnlimited.Models;
using PartsUnlimited.ProductSearch;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PartsUnlimited.ProductSearch.Tests
{
    [TestClass()]
    public class StringContainsProductSearchTests
    {
        private static readonly IEnumerable<string> s_productTitles = new[] { "word in the middle", "something", "something outside", "inside where outside" };

        [TestMethod()]
        
        public async Task SearchSuccess()
        {
            var productList = s_productTitles.Select(o => new Product { Title = o }).ToList();
            var context = Substitute.For<IPartsUnlimitedContext>();
            var productDbSet = productList.ToDbSet();

            context.Products.Returns(productDbSet);

            var searcher = new StringContainsProductSearch(context);

            var thing = await searcher.Search("thing");

            Assert.AreEqual(new string[] { "something", "something outside" }, thing.Select(o => o.Title));
        }
    }
}