using ClosedXML.Excel;
using crowlerSj.Db;
using crowlerSj.signalR;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private readonly SearchContext _context;
        private bool _isCrawlerRunning = false;
        private static CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();


        public HomeController(IServiceScopeFactory scopeFactory, ILogger<HomeController> logger, SearchContext context)
        {
            _scopeFactory = scopeFactory;
            _logger = logger;
            _context = context;
        }

        public IActionResult Index()
        {
            var setting = _context.Settings.ToList().FirstOrDefault();
            _isCrawlerRunning = setting.IsCrowl;
            ViewBag.IsCrawlerRunning = _isCrawlerRunning;
            return View();
        }


        [HttpPost]
        public IActionResult Start(string search, long crowlId)
        {
            _logger.LogInformation("Index method called. Starting crawler.");
            _isCrawlerRunning = true;

            var setting = _context.Settings.FirstOrDefault();
            if (setting != null)
            {
                setting.IsCrowl = _isCrawlerRunning;
                _context.SaveChanges();
            }

            if (crowlId ==0 )
            {
                var crowl = _context.Crowls.Add(new Crowl { });
                crowlId = crowl.Entity.Id;
            }

            _cancellationTokenSource = new CancellationTokenSource();
            Task.Run(() => Crawl(search, crowlId, _cancellationTokenSource.Token), _cancellationTokenSource.Token);
            return Redirect("~/home/index");
        }

        [HttpPost]
        public IActionResult Stop()
        {
            if (_isCrawlerRunning)
            {
                _cancellationTokenSource.Cancel();
                _isCrawlerRunning = false;
                ViewBag.IsCrawlerRunning = false;
            }
            try
            {
                var chromeDriverProcesses = Process.GetProcessesByName("chromedriver");
                foreach (var process in chromeDriverProcesses)
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error while closing ChromeDriver processes: {ex.Message}");
            }


            var setting = _context.Settings.FirstOrDefault();
            if (setting != null)
            {
                setting.IsCrowl = _isCrawlerRunning;
                _context.SaveChanges();
            }

            // اینجا به کلاینت‌ها پیغام ارسال می‌کنیم
            var hubContext = (IHubContext<CrawlerHub>)HttpContext.RequestServices.GetService(typeof(IHubContext<CrawlerHub>));
            hubContext.Clients.All.SendAsync("ReceiveMessage", "Crawler stopped");

            return RedirectToAction("Index");
        }
        private void Crawl(string search, long crowlId, CancellationToken token)
        {
            int retryLimit = 3;
            int retries = 0;

            while (!token.IsCancellationRequested && _isCrawlerRunning && retries < retryLimit)
            {
                _logger.LogInformation("Crawl method started.");

                var queries = GetSearchKeywords(search);
                var options = new ChromeOptions();
                options.AddArgument("--disable-gpu");

                using (IWebDriver driver = new ChromeDriver(options))
                {
                    int currentPage = 1;
                    try
                    {
                        foreach (var query in queries)
                        {
                            if (token.IsCancellationRequested)
                            {
                                _logger.LogInformation("Crawl operation canceled before processing.");
                                return;
                            }

                            // Navigate and process each query result
                            NavigateAndProcessResults(driver, query, ref currentPage, crowlId);

                            if (token.IsCancellationRequested)
                            {
                                _logger.LogInformation("Crawl operation canceled after processing.");
                                return;
                            }
                        }

                        retries = 0;
                    }
                    catch (WebDriverException ex)
                    {
                        _logger.LogError($"WebDriverException caught on page {currentPage}: {ex.Message}");
                        HandleError(ex, currentPage);
                        retries++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Exception caught on page {currentPage}: {ex.Message}");
                        HandleError(ex, currentPage);
                        retries++;
                    }
                    finally
                    {
                        driver.Quit(); // Ensure all resources are freed


                        try
                        {
                            var chromeTabs = Process.GetProcessesByName("chrome").Where(p => p.MainWindowTitle.Contains("data:") || p.MainWindowTitle.Contains("about:blank"));
                            foreach (var tab in chromeTabs)
                            {
                                tab.Kill();
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"Error while closing Chrome tabs: {ex.Message}");
                        }


                    }
                }

                if (retries >= retryLimit)
                {
                    _logger.LogError("Retry limit reached. Stopping crawler.");
                    _isCrawlerRunning = false;
                }
                else if (_isCrawlerRunning)
                {
                    _logger.LogInformation("Retrying crawl...");
                    Thread.Sleep(200); // Adjust according to need
                }
            }

            _logger.LogInformation("Crawl method finished.");
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
        private void ProcessCurrentPageResults(IWebDriver driver, int currentPage, long crowlId)
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
                    searchResult.CrowlId = crowlId;

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
     
        private void HandleError(Exception ex, int currentPage)
        {
            SaveErrorLog(ex, currentPage);
            // اعمال دیگری که ممکن است برای مدیریت خطاها مورد نیاز باشد
        }

        private void NavigateAndProcessResults(IWebDriver driver, string query, ref int currentPage, long crowlId)
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
                ProcessCurrentPageResults(driver, currentPage,  crowlId);

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
     
        private void SaveErrorLog(Exception ex, int currentPage)
        {
            // مثال ساده برای ذخیره لاگ خطاها
            _logger.LogError($"Error on page {currentPage}: {ex.Message}");
            // می‌توانید لاگ خطا را در دیتابیس یا فایل ذخیره کنید
        }

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

      

        public IActionResult List()
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                var reports = context.Crowls.ToList();
                _logger.LogInformation($"{reports.Count} items retrieved from database for reports.");

              
                return View(reports);
            }
        }

        public IActionResult Reports(long crowlId)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                var reports = context.SearchResults
                    .Where(r => r.CrowlId == crowlId) // اعمال شرط تاریخ
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
    }
}
