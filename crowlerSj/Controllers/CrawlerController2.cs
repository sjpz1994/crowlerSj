using Microsoft.AspNetCore.Mvc;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

public class CrawlerController2 : Controller
{
    public IActionResult Index()
    {
        string query = "دوربین مدار بسته خرید فروش تعمیرات نصب";
        var results = new List<SearchResult>();

        var options = new ChromeOptions();
        options.AddArgument("--headless"); // اجرای مرورگر به صورت مخفی
        options.AddArgument("--disable-gpu");

        using (var driver = new ChromeDriver(options))
        {
            driver.Navigate().GoToUrl($"https://www.google.com/search?q={query}");
            Thread.Sleep(2000); // تاخیر برای بارگذاری صفحه

            var searchResults = driver.FindElements(By.CssSelector(".tF2Cxc"));

            foreach (var result in searchResults)
            {
                var titleElement = result.FindElement(By.CssSelector(".DKV0Md"));
                var linkElement = result.FindElement(By.CssSelector(".yuRUbf a"));
                var snippetElement = result.FindElement(By.CssSelector(".IsZvec"));

                var searchResult = new SearchResult
                {
                    Title = titleElement.Text,
                    Link = linkElement.GetAttribute("href"),
                    Snippet = snippetElement.Text
                };

                // باز کردن لینک و جمع‌آوری اطلاعات تماس و دسته‌بندی
                driver.Navigate().GoToUrl(searchResult.Link);
                Thread.Sleep(2000); // تاخیر برای بارگذاری صفحه

                searchResult.Phone = ExtractPhoneNumber(driver.PageSource);
                searchResult.Address = ExtractAddress(driver.PageSource);
                searchResult.Category = ExtractCategory(driver.PageSource);

                results.Add(searchResult);
                driver.Navigate().Back(); // بازگشت به صفحه نتایج جستجو
                Thread.Sleep(2000);
            }
        }

        return View(results);
    }

    private string ExtractPhoneNumber(string html)
    {
        var phoneMatch = Regex.Match(html, @"\(?\d{3}\)?-?\s*\d{3}-\d{4}");
        return phoneMatch.Success ? phoneMatch.Value : "Not Found";
    }

    private string ExtractAddress(string html)
    {
        var addressMatch = Regex.Match(html, @"\d{1,5}\s\w+\s\w+");
        return addressMatch.Success ? addressMatch.Value : "Not Found";
    }

    private string ExtractCategory(string html)
    {
        if (html.Contains("فروش"))
        {
            return "فروش";
        }
        else if (html.Contains("تعمیرات"))
        {
            return "تعمیرات";
        }
        else if (html.Contains("نصب"))
        {
            return "نصب";
        }
        else
        {
            return "نامشخص";
        }
    }
}

