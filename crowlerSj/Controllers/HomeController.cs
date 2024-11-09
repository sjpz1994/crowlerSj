using ClosedXML.Excel;
using crowlerSj.Db;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace CrowlerSj.Controllers
{
    public class HomeController : Controller
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly ILogger<HomeController> _logger;

        public HomeController(IServiceScopeFactory scopeFactory, ILogger<HomeController> logger)
        {
            _scopeFactory = scopeFactory;
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult start(string search)
        {
            _logger.LogInformation("Index method called. Starting crawler.");
            Task.Run(() => Crawl(search));
            return Json("ربات درحال اجرا می باشد");
        }

        private List<string> GetSearchKeywords(string search)
        {
            return new List<string>
            {
                search
            };
        }

        private string ExtractPhoneNumber(string html)
        {
            var phoneMatches = Regex.Matches(html, @"((0?9)|(\+?989))\d{2}\W?\d{3}\W?\d{4}|^0\d{2,3}-\d{8}$");
            if (phoneMatches.Count > 0)
            {
                StringBuilder phoneNumberBuilder = new StringBuilder();
                foreach (Match match in phoneMatches)
                {
                    phoneNumberBuilder.Append(match.Value);
                    phoneNumberBuilder.Append(", "); // اضافه کردن جداکننده بین شماره‌ها
                }
                phoneNumberBuilder.Remove(phoneNumberBuilder.Length - 2, 2); // حذف آخرین کاراکتر اضافه شده (جداکننده اضافی)
                return phoneNumberBuilder.ToString();
            }
            else
            {
                return "Not Found";
            }
        }

        private bool IsLinkAlreadyProcessed(string link)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                return context.SearchResults.Any(r => r.Link == link);
            }
        }
        private void Crawl(string search)
        {
            _logger.LogInformation("Crawl method started.");
            var queries = GetSearchKeywords(search);
            var options = new ChromeOptions();
            options.AddArgument("--disable-gpu");

            IWebDriver driver = null;
            int currentPage = 1;

            try
            {
                driver = new ChromeDriver(options);
                foreach (var query in queries)
                {
                    NavigateAndProcessResults(driver, query, ref currentPage);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Exception caught on page {currentPage}: {ex.Message}");
                SaveErrorLog(ex, currentPage);
            }
            finally
            {
                driver?.Quit();
            }
            _logger.LogInformation("Crawl method finished.");
        }

        /* private void Crawl()
         {
             _logger.LogInformation("Crawl method started.");
             var keywords = GetSearchKeywords();
             var options = new ChromeOptions();
             options.AddArgument("--disable-gpu");

             IWebDriver driver = null;

             try
             {
                 driver = new ChromeDriver(options);

                 foreach (var keyword in keywords)
                 {
                     int currentPage = 1;
                     NavigateAndProcessResults(driver, keyword, ref currentPage);
                 }
             }
             catch (Exception ex)
             {
                 _logger.LogError($"Exception caught: {ex.Message}");
                 SaveErrorLog(ex);
             }
             finally
             {
                 driver?.Quit();
             }
             _logger.LogInformation("Crawl method finished.");
         }*/

        /* private void NavigateAndProcessResults(IWebDriver driver, string query, ref int currentPage)
         {
             _logger.LogInformation($"Navigating to Google search with query: {query}");
             driver.Navigate().GoToUrl($"https://www.google.com/search?q={query}");
             Thread.Sleep(1000);

             if (NeedManualIntervention(driver))
             {
                 _logger.LogInformation("Manual intervention needed. Please solve the CAPTCHA or accept the cookies in your browser and then press Enter to continue...");
                 Console.ReadLine();
             }

             while (true) // Infinite loop to process all search results
             {
                 ProcessCurrentPageResults(driver, currentPage);

                 var nextPageElement = driver.FindElements(By.CssSelector("#pnnext")).FirstOrDefault();
                 if (nextPageElement == null)
                 {
                     // No more pages, break out of the loop
                     break;
                 }

                 _logger.LogInformation($"Navigating to next page: {currentPage + 1}");
                 nextPageElement.Click();
                 Thread.Sleep(2000); // Wait for the next page to load
                 currentPage++;

                 if (NeedManualIntervention(driver))
                 {
                     _logger.LogInformation("Manual intervention needed. Please solve the CAPTCHA or accept the cookies in your browser and then press Enter to continue...");
                     Console.ReadLine();
                 }
             }
         }*/
        private void NavigateAndProcessResults(IWebDriver driver, string query, ref int currentPage)
        {
            _logger.LogInformation($"Navigating to Bing search with query: {query}");
            driver.Navigate().GoToUrl($"https://www.bing.com/search?q={query}");
            Thread.Sleep(1000);

            if (NeedManualIntervention(driver))
            {
                _logger.LogInformation("Manual intervention needed. Please solve the CAPTCHA or accept the cookies in your browser and then press Enter to continue...");
                Console.ReadLine();
            }

            while (true) // Infinite loop to process all search results
            {
                ProcessCurrentPageResults(driver, currentPage);

                var nextPageElement = driver.FindElements(By.CssSelector(".sb_pagN")).FirstOrDefault();
                if (nextPageElement == null)
                {
                    // No more pages, break out of the loop
                    break;
                }

                _logger.LogInformation($"Navigating to next page: {currentPage + 1}");
                nextPageElement.Click();
                Thread.Sleep(2000); // Wait for the next page to load
                currentPage++;

                if (NeedManualIntervention(driver))
                {
                    _logger.LogInformation("Manual intervention needed. Please solve the CAPTCHA or accept the cookies in your browser and then press Enter to continue...");
                    Console.ReadLine();
                }
            }
        }
        private void ProcessCurrentPageResults(IWebDriver driver, int currentPage)
        {
            var searchResults = driver.FindElements(By.CssSelector(".b_algo"));

            if (searchResults.Count == 0)
            {
                return; // No results to process
            }

            foreach (var result in searchResults)
            {
                try
                {
                    _logger.LogInformation("Processing search result.");
                    var titleElement = result.FindElement(By.CssSelector("h2 a"));
                    var linkElement = result.FindElement(By.CssSelector("h2 a"));

                    var searchResult = new SearchResult
                    {
                        Title = titleElement.Text,
                        Link = linkElement.GetAttribute("href"),
                        Snippet = result.Text
                    };

                    _logger.LogInformation($"Title: {searchResult.Title}, Link: {searchResult.Link}");

                    // Check if the link has been processed before
                    if (IsLinkAlreadyProcessed(searchResult.Link))
                    {
                        _logger.LogInformation($"Link '{searchResult.Link}' has already been processed. Skipping...");
                        continue;
                    }

                    // Open the link in a new tab
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    driver.Navigate().GoToUrl(searchResult.Link);
                    Thread.Sleep(1000);

                    // Check if the page load time is acceptable
                    if (!IsPageLoadTimeAcceptable(driver, 3))
                    {
                        _logger.LogInformation($"Page load time for link '{searchResult.Link}' exceeds 3 seconds. Skipping...");
                        driver.Close();
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                        continue;
                    }

                    // Extract phone number
                    searchResult.Phone = ExtractPhoneNumber(driver.PageSource);
                    searchResult.Address = ExtractAddress(driver.PageSource);
                    searchResult.Category = ExtractCategory(driver.PageSource);

                    // Save the result to the database
                    SaveToDatabase(searchResult);

                    // Close the tab and switch back to the main tab
                    driver.Close();
                    driver.SwitchTo().Window(driver.WindowHandles.First());
                }
                catch (NoSuchElementException ex)
                {
                    _logger.LogWarning($"NoSuchElementException caught on page {currentPage}: {ex.Message}");
                    SaveErrorLog(ex, currentPage);
                    driver.SwitchTo().Window(driver.WindowHandles.First());
                    continue;
                }
                catch (WebDriverException ex)
                {
                    _logger.LogError($"WebDriverException caught on page {currentPage}: {ex.Message}");
                    SaveErrorLog(ex, currentPage);
                    driver.SwitchTo().Window(driver.WindowHandles.First());
                    continue;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Exception caught on page {currentPage}: {ex.Message}");
                    SaveErrorLog(ex, currentPage);
                    driver.SwitchTo().Window(driver.WindowHandles.First());
                    continue;
                }
            }
        }
        private void SaveErrorLog(Exception ex, int currentPage)
        {
            // مثال ساده برای ذخیره لاگ خطاها
            _logger.LogError($"Error on page {currentPage}: {ex.Message}");
            // می‌توانید لاگ خطا را در دیتابیس یا فایل ذخیره کنید
        }

        /*
                private void ProcessCurrentPageResults(IWebDriver driver, int currentPage)
                {
                    var searchResults = driver.FindElements(By.CssSelector(".g"));

                    if (searchResults.Count == 0)
                    {
                        return; // No results to process
                    }

                    foreach (var result in searchResults)
                    {
                        try
                        {
                            _logger.LogInformation("Processing search result.");
                            var titleElement = result.FindElement(By.CssSelector("h3"));
                            var linkElement = result.FindElement(By.CssSelector("a"));

                            var searchResult = new SearchResult
                            {
                                Title = titleElement.Text,
                                Link = linkElement.GetAttribute("href"),
                                Snippet = result.Text
                            };

                            _logger.LogInformation($"Title: {searchResult.Title}, Link: {searchResult.Link}");

                            // Check if the link has been processed before
                            if (IsLinkAlreadyProcessed(searchResult.Link))
                            {
                                _logger.LogInformation($"Link '{searchResult.Link}' has already been processed. Skipping...");
                                continue;
                            }

                            // Open the link in a new tab
                            ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
                            driver.SwitchTo().Window(driver.WindowHandles.Last());
                            driver.Navigate().GoToUrl(searchResult.Link);
                            Thread.Sleep(1000);

                            // Check if the page load time is acceptable
                            if (!IsPageLoadTimeAcceptable(driver, 3))
                            {
                                _logger.LogInformation($"Page load time for link '{searchResult.Link}' exceeds 3 seconds. Skipping...");
                                driver.Close();
                                driver.SwitchTo().Window(driver.WindowHandles.First());
                                continue;
                            }

                            // Extract phone number
                            searchResult.Phone = ExtractPhoneNumber(driver.PageSource);
                            searchResult.Address = ExtractAddress(driver.PageSource);
                            searchResult.Category = ExtractCategory(driver.PageSource);

                            // Save the result to the database
                            SaveToDatabase(searchResult);

                            // Close the tab and switch back to the main tab
                            driver.Close();
                            driver.SwitchTo().Window(driver.WindowHandles.First());
                        }
                        catch (NoSuchElementException ex)
                        {
                            _logger.LogWarning($"NoSuchElementException caught on page {currentPage}: {ex.Message}");
                            SaveErrorLog(ex);
                            driver.SwitchTo().Window(driver.WindowHandles.First());
                            continue;
                        }
                        catch (WebDriverException ex)
                        {
                            _logger.LogError($"WebDriverException caught on page {currentPage}: {ex.Message}");
                            SaveErrorLog(ex);
                            driver.SwitchTo().Window(driver.WindowHandles.First());
                            continue;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"Exception caught on page {currentPage}: {ex.Message}");
                            SaveErrorLog(ex);
                            driver.SwitchTo().Window(driver.WindowHandles.First());
                            continue;
                        }
                    }
                }*/

        private bool IsPageLoadTimeAcceptable(IWebDriver driver, int maxLoadTimeInSeconds)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxLoadTimeInSeconds));
                wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
                return true;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }

        private void SaveErrorLog(Exception ex)
        {
            // Save error details to a file or database for further analysis
            _logger.LogError($"Error: {ex}");
        }

        [HttpGet]
        public IActionResult DownloadExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Search Results");
                var currentRow = 1;

                // هدرها
                worksheet.Cell(currentRow, 1).Value = "Title";
                worksheet.Cell(currentRow, 2).Value = "Snippet";
                worksheet.Cell(currentRow, 3).Value = "Phone";
                worksheet.Cell(currentRow, 4).Value = "Address";
                worksheet.Cell(currentRow, 5).Value = "Category";
                worksheet.Cell(currentRow, 6).Value = "Link";

                // داده‌های جستجو
                using (var scope = _scopeFactory.CreateScope())
                {
                    var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                    var searchResults = context.SearchResults.ToList();

                    foreach (var result in searchResults)
                    {
                        currentRow++;
                        worksheet.Cell(currentRow, 1).Value = result.Title;
                        worksheet.Cell(currentRow, 2).Value = result.Snippet;
                        worksheet.Cell(currentRow, 3).Value = result.Phone;
                        worksheet.Cell(currentRow, 4).Value = result.Address;
                        worksheet.Cell(currentRow, 5).Value = result.Category;
                        worksheet.Cell(currentRow, 6).Value = result.Link;
                    }
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Reports.xlsx");
                }
            }
        }

        private bool NeedManualIntervention(IWebDriver driver)
        {
            try
            {
                var captcha = driver.FindElement(By.CssSelector("div#recaptcha"));
                if (captcha.Displayed)
                {
                    return true;
                }
            }
            catch (NoSuchElementException) { }

            try
            {
                var cookieMessage = driver.FindElement(By.CssSelector("div#cookieConsent"));
                if (cookieMessage.Displayed)
                {
                    return true;
                }
            }
            catch (NoSuchElementException) { }

            return false;
        }

        private string ExtractAddress(string html)
        {
            var addressMatch = Regex.Match(html, @"\d{1,5}\s\w+\s\w+");
            return addressMatch.Success ? addressMatch.Value : "Not Found";
        }

        private string ExtractEmail(string html)
        {
            var addressMatch = Regex.Match(html, @"([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+");
            return addressMatch.Success ? addressMatch.Value : "Not Found";
        }

        private string ExtractCategory(string html)
        {
            string result = "";
            if (html.Contains("فروش"))
            {
                result += "+فروش";
            }
            if (html.Contains("تعمیرات"))
            {
                result += "+تعمیرات";
            }
            if (html.Contains("نصب"))
            {
                result += "+نصب";
            }
            if (string.IsNullOrEmpty(result))
            {
                result = "نامشخص";
            }
            return result;
        }

        private void SaveToDatabase(SearchResult searchResult)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                context.SearchResults.Add(searchResult);
                context.SaveChanges();
                _logger.LogInformation($"Search result saved to database: {searchResult.Link}");
            }
        }

        public IActionResult List()
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                var reports = context.SearchResults.ToList();
                _logger.LogInformation($"{reports.Count} items retrieved from database for reports.");

                var list = new List<DateTime>();
                foreach (var item in reports)
                {
                    // Check if the date (ignoring time) already exists in the list
                    if (!list.Exists(x => x.Date == item.InsertTime.Date))
                    {
                        list.Add(item.InsertTime);
                    }
                }
                return View(list);
            }
        }

        public IActionResult Reports(DateTime date)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                var reports = context.SearchResults
                    .Where(r => r.InsertTime.Date == date.Date) // اعمال شرط تاریخ
                    .ToList();

                if (reports.Any())
                {
                    _logger.LogInformation($"{reports.Count} items retrieved for the selected date.");
                }
                else
                {
                    _logger.LogWarning("No reports found for the selected date.");
                }

                return View(reports); // در صورت نیاز ویوی دیگری برگردانده شود
            }
        }

    }
}
