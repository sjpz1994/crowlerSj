using ClosedXML.Excel;
using crowlerSj.Db;
using crowlerSj.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
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
        private CancellationTokenSource _cancellationTokenSource;
        private static readonly object _crawlerLock = new object();

        public HomeController(IServiceScopeFactory scopeFactory, ILogger<HomeController> logger, SearchContext context)
        {
            _scopeFactory = scopeFactory;
            _logger = logger;
            _context = context;
        }

        public IActionResult Index()
        {
            try
            {
                using (var scope = _scopeFactory.CreateScope())
                {
                    var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                    // فقط برای اطمینان از وجود تنظیمات
                    var setting = context.Settings.FirstOrDefault();
                    if (setting == null)
                    {
                        setting = new Setting { Id = 1, IsCrowl = false };
                        context.Settings.Add(setting);
                        context.SaveChanges();
                        _logger.LogInformation("جدول Settings خالی بود. ردیف جدید ایجاد شد.");
                    }

                    // لود لاگ‌ها با بهینه‌سازی
                    ViewBag.Logs = context.Logs
                        .OrderByDescending(l => l.Timestamp)
                        .Take(50)
                        .AsNoTracking()
                        .ToList();
                }

                return View();
            }
            catch (Exception ex)
            {
                _logger.LogError($"خطا در بارگذاری صفحه Index: {ex.Message}");
                SaveLog($"خطا در بارگذاری صفحه Index: {ex.Message}", "Error");
                ViewBag.ErrorMessage = $"خطا در بارگذاری صفحه: {ex.Message}";
                return View();
            }
        }

        [HttpPost]
        public IActionResult Start(string search, long crowlId)
        {
            lock (_crawlerLock)
            {
                if (_isCrawlerRunning)
                {
                    _logger.LogWarning("کرولر در حال اجراست. لطفاً برنامه را متوقف کنید.");
                    SaveLog("کرولر در حال اجراست. لطفاً برنامه را متوقف کنید.", "Warning");
                    return RedirectToAction("Index");
                }

                CleanupProcesses();

                _logger.LogInformation("شروع کرولر با کلمه جستجو: {search}", search);

                long validCrowlId;
                int startPage = 1;

                using (var scope = _scopeFactory.CreateScope())
                {
                    var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                    try
                    {
                        if (crowlId > 0 && context.Crowls.Any(c => c.Id == crowlId))
                        {
                            var crowl = context.Crowls.FirstOrDefault(c => c.Id == crowlId);
                            startPage = crowl.CurrentPage;
                            validCrowlId = crowlId;
                            search = crowl.Title; // استفاده از عنوان کرول موجود
                            _logger.LogInformation($"ادامه کرول با Id: {validCrowlId}, صفحه: {startPage}");
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(search))
                            {
                                _logger.LogError("کلمه جستجو خالی است.");
                                SaveLog("کلمه جستجو خالی است.", "Error");
                                return StatusCode(400, "کلمه جستجو نمی‌تواند خالی باشد.");
                            }
                            var crowl = new Crowl { Title = search, CurrentPage = 1, InsertTime = DateTime.Now };
                            context.Crowls.Add(crowl);
                            context.SaveChanges();
                            validCrowlId = crowl.Id;
                            _logger.LogInformation($"کرول جدید ساخته شد با Id: {validCrowlId}");
                        }

                        SaveLog($"شروع کرولر با کلمه جستجو: {search}", "Information", validCrowlId);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"خطا در ایجاد کرول: {ex.Message}");
                        SaveLog($"خطا در ایجاد کرول: {ex.Message}", "Error");
                        return StatusCode(500, $"خطا در شروع کرولر: {ex.Message}");
                    }
                }

                _isCrawlerRunning = true;
                using (var scope = _scopeFactory.CreateScope())
                {
                    var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                    try
                    {
                        var setting = context.Settings.FirstOrDefault();
                        if (setting == null)
                        {
                            setting = new Setting { Id = 1, IsCrowl = true };
                            context.Settings.Add(setting);
                        }
                        else
                        {
                            setting.IsCrowl = true;
                        }
                        context.SaveChanges();
                        _logger.LogInformation($"تنظیمات دیتابیس آپدیت شد: IsCrowl = true برای کرول Id {validCrowlId}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"خطا در آپدیت تنظیمات: {ex.Message}");
                        SaveLog($"خطا در آپدیت تنظیمات: {ex.Message}", "Error");
                    }
                }

                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = new CancellationTokenSource();
                Task.Run(async () =>
                {
                    try
                    {
                        await CrawlAsync(search, validCrowlId, startPage, _cancellationTokenSource.Token);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"خطا در اجرای کرول: {ex.Message}");
                        SaveLog($"خطا در اجرای کرول: {ex.Message}", "Error", validCrowlId);
                        ResetCrawlerState(validCrowlId);
                    }
                }, _cancellationTokenSource.Token);

                return RedirectToAction("Index");
            }
        }

        private void CleanupProcesses()
        {
            foreach (var process in Process.GetProcessesByName("chromedriver"))
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(2000);
                    _logger.LogInformation($"پروسه chromedriver با PID {process.Id} بسته شد.");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"خطا در بستن پروسه chromedriver: {ex.Message}");
                }
            }

            foreach (var process in Process.GetProcessesByName("chrome"))
            {
                if (process.MainWindowTitle.Contains("Bing"))
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(2000);
                        _logger.LogInformation($"پروسه chrome با PID {process.Id} بسته شد.");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"خطا در بستن پروسه chrome: {ex.Message}");
                    }
                }
            }
        }

        private async Task CrawlAsync(string search, long crowlId, int startPage, CancellationToken token)
        {
            const int retryLimit = 3;
            int retries = 0;

            while (!token.IsCancellationRequested && _isCrawlerRunning && retries < retryLimit)
            {
                _logger.LogInformation("شروع عملیات کرول.");
                SaveLog("شروع عملیات کرول.", "Information", crowlId);

                var queries = GetSearchKeywords(search);
                var options = new ChromeOptions();
                options.AddArgument("--disable-gpu");

                try
                {
                    using (IWebDriver driver = new ChromeDriver(options))
                    {
                        int currentPage = startPage;
                        try
                        {
                            foreach (var query in queries)
                            {
                                if (token.IsCancellationRequested)
                                {
                                    _logger.LogInformation("عملیات کرول لغو شد.");
                                    SaveLog("عملیات کرول لغو شد.", "Information", crowlId);
                                    return;
                                }

                                currentPage = await NavigateAndProcessResultsAsync(driver, query, currentPage, crowlId, token);
                            }

                            retries = 0;
                        }
                        catch (WebDriverException ex)
                        {
                            _logger.LogError($"خطای WebDriver در صفحه {currentPage}: {ex.Message}");
                            SaveLog($"خطای WebDriver در صفحه {currentPage}: {ex.Message}", "Error", crowlId);
                            retries++;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"خطای عمومی در صفحه {currentPage}: {ex.Message}");
                            SaveLog($"خطای عمومی در صفحه {currentPage}: {ex.Message}", "Error", crowlId);
                            retries++;
                        }
                        finally
                        {
                            try
                            {
                                driver.Quit();
                                _logger.LogInformation("درایور Selenium بسته شد.");
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning($"خطا در بستن درایور: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"خطا در ایجاد درایور Chrome: {ex.Message}");
                    SaveLog($"خطا در ایجاد درایور Chrome: {ex.Message}", "Error", crowlId);
                    retries++;
                }

                if (retries >= retryLimit)
                {
                    _logger.LogError("حداکثر تلاش‌ها انجام شد. توقف کرولر.");
                    SaveLog("حداکثر تلاش‌ها انجام شد. توقف کرولر.", "Error", crowlId);
                    ResetCrawlerState(crowlId);
                }
                else if (_isCrawlerRunning)
                {
                    _logger.LogInformation("تلاش مجدد برای کرول...");
                    SaveLog("تلاش مجدد برای کرول...", "Information", crowlId);
                    await Task.Delay(1000, token);
                }
            }

            ResetCrawlerState(crowlId);
        }

        private void ResetCrawlerState(long? crowlId = null)
        {
            lock (_crawlerLock)
            {
                _isCrawlerRunning = false;
                using (var scope = _scopeFactory.CreateScope())
                {
                    var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                    var setting = context.Settings.FirstOrDefault();
                    if (setting != null)
                    {
                        setting.IsCrowl = false;
                        try
                        {
                            context.SaveChanges();
                            _logger.LogInformation("تنظیمات دیتابیس ریست شد: IsCrowl = false");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"خطا در ریست تنظیمات دیتابیس: {ex.Message}");
                            SaveLog($"خطا در ریست تنظیمات دیتابیس: {ex.Message}", "Error");
                        }
                    }
                    else
                    {
                        setting = new Setting { Id = 1, IsCrowl = false };
                        context.Settings.Add(setting);
                        context.SaveChanges();
                        _logger.LogInformation("جدول Settings خالی بود. ردیف جدید ایجاد شد: IsCrowl = false");
                    }
                }
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                if (crowlId.HasValue)
                {
                    _logger.LogInformation("عملیات کرول به پایان رسید.");
                    SaveLog("عملیات کرول به پایان رسید.", "Information", crowlId);
                }
            }
        }

        private void SaveToDatabase(SearchResult searchResult)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                try
                {
                    context.SearchResults.Add(searchResult);
                    context.SaveChanges();
                    _logger.LogInformation($"نتیجه جستجو ذخیره شد: {searchResult.Link}");
                    SaveLog($"نتیجه جستجو ذخیره شد: {searchResult.Link}", "Information", searchResult.CrowlId);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"خطا در ذخیره نتیجه جستجو: {ex.Message}");
                    SaveLog($"خطا در ذخیره نتیجه جستجو: {ex.Message}", "Error", searchResult.CrowlId);
                }
            }
        }

        private async Task ProcessCurrentPageResultsAsync(IWebDriver driver, int currentPage, long crowlId, CancellationToken token)
        {
            var searchResults = driver.FindElements(By.CssSelector(".b_algo"));
            if (!searchResults.Any())
            {
                _logger.LogInformation("هیچ نتیجه‌ای در صفحه یافت نشد.");
                SaveLog("هیچ نتیجه‌ای در صفحه یافت نشد.", "Information", crowlId);
                return;
            }

            foreach (var result in searchResults)
            {
                if (token.IsCancellationRequested)
                {
                    return;
                }

                try
                {
                    var titleElement = result.FindElement(By.CssSelector("h2 a"));
                    var linkElement = result.FindElement(By.CssSelector("h2 a"));

                    var searchResult = new SearchResult
                    {
                        Title = titleElement.Text,
                        Link = linkElement.GetAttribute("href"),
                        Snippet = result.Text,
                        CrowlId = crowlId
                    };

                    if (IsLinkAlreadyProcessed(searchResult.Link))
                    {
                        _logger.LogInformation($"لینک '{searchResult.Link}' قبلاً پردازش شده. رد شدن...");
                        SaveLog($"لینک '{searchResult.Link}' قبلاً پردازش شده. رد شدن...", "Information", crowlId);
                        continue;
                    }

                    ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    driver.Navigate().GoToUrl(searchResult.Link);
                    await Task.Delay(1000, token);

                    if (!IsPageLoadTimeAcceptable(driver, 5))
                    {
                        _logger.LogInformation($"زمان لود صفحه '{searchResult.Link}' بیش از 5 ثانیه است. رد شدن...");
                        SaveLog($"زمان لود صفحه '{searchResult.Link}' بیش از 5 ثانیه است. رد شدن...", "Information", crowlId);
                        driver.Close();
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                        continue;
                    }

                    searchResult.Phone = ExtractPhoneNumber(driver.PageSource);
                    searchResult.Address = ExtractAddress(driver.PageSource);
                    searchResult.Category = ExtractCategory(driver.PageSource);

                    SaveToDatabase(searchResult);

                    driver.Close();
                    driver.SwitchTo().Window(driver.WindowHandles.First());
                }
                catch (Exception ex)
                {
                    _logger.LogError($"خطا در پردازش نتیجه صفحه {currentPage}: {ex.Message}");
                    SaveLog($"خطا در پردازش نتیجه صفحه {currentPage}: {ex.Message}", "Error", crowlId);
                    if (driver.WindowHandles.Count > 1)
                    {
                        driver.Close();
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                    }
                }
            }

            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                var crowl = context.Crowls.FirstOrDefault(c => c.Id == crowlId);
                if (crowl != null)
                {
                    crowl.CurrentPage = currentPage;
                    try
                    {
                        context.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"خطا در آپدیت صفحه کرول: {ex.Message}");
                        SaveLog($"خطا در آپدیت صفحه کرول: {ex.Message}", "Error", crowlId);
                    }
                }
            }
        }

        private async Task<int> NavigateAndProcessResultsAsync(IWebDriver driver, string query, int currentPage, long crowlId, CancellationToken token)
        {
            _logger.LogInformation($"رفتن به جستجوی بینگ با کلمه: {query}");
            SaveLog($"رفتن به جستجوی بینگ با کلمه: {query}", "Information", crowlId);
            driver.Navigate().GoToUrl($"https://www.bing.com/search?q={Uri.EscapeDataString(query)}&first={(currentPage - 1) * 10}");
            await Task.Delay(1000, token);

            if (HandleConsentOrCaptcha(driver))
            {
                _logger.LogInformation("کپچا شناسایی شد. نیاز به دخالت دستی.");
                SaveLog("کپچا شناسایی شد. نیاز به دخالت دستی.", "Warning", crowlId);
                Console.ReadLine();
            }

            int maxPageAttempts = 3;
            int pageAttempts = 0;

            while (true)
            {
                if (token.IsCancellationRequested)
                {
                    return currentPage;
                }

                await ProcessCurrentPageResultsAsync(driver, currentPage, crowlId, token);

                try
                {
                    var nextPageElement = driver.FindElements(By.CssSelector(".sb_pagN")).FirstOrDefault()
                        ?? driver.FindElements(By.CssSelector("a[title='Next page']")).FirstOrDefault();
                    if (nextPageElement == null || pageAttempts >= maxPageAttempts)
                    {
                        _logger.LogInformation("صفحه بعدی وجود ندارد یا تلاش‌ها تمام شد.");
                        SaveLog("صفحه بعدی وجود ندارد یا تلاش‌ها تمام شد.", "Information", crowlId);
                        break;
                    }

                    _logger.LogInformation($"رفتن به صفحه بعدی: {currentPage + 1}");
                    SaveLog($"رفتن به صفحه بعدی: {currentPage + 1}", "Information", crowlId);
                    nextPageElement.Click();
                    await Task.Delay(new Random().Next(2000, 4000), token);
                    currentPage++;
                    pageAttempts = 0;

                    if (HandleConsentOrCaptcha(driver))
                    {
                        _logger.LogInformation("کپچا شناسایی شد. نیاز به دخالت دستی.");
                        SaveLog("کپچا شناسایی شد. نیاز به دخالت دستی.", "Warning", crowlId);
                        Console.ReadLine();
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"خطا در رفتن به صفحه بعدی: {ex.Message}");
                    SaveLog($"خطا در رفتن به صفحه بعدی: {ex.Message}", "Error", crowlId);
                    pageAttempts++;
                    await Task.Delay(2000, token);
                }
            }

            return currentPage;
        }

        private List<string> GetSearchKeywords(string search)
        {
            return new List<string> { search };
        }

        private string ExtractPhoneNumber(string html)
        {
            var phoneMatches = Regex.Matches(html, @"((0?9)|(\+?989))\d{2}\W?\d{3}\W?\d{4}|^0\d{2,3}-\d{8}$");
            if (phoneMatches.Count > 0)
            {
                var phoneNumberBuilder = new StringBuilder();
                foreach (Match match in phoneMatches)
                {
                    phoneNumberBuilder.Append(match.Value + ", ");
                }
                return phoneNumberBuilder.ToString().TrimEnd(',', ' ');
            }
            return "Not Found";
        }

        private bool IsLinkAlreadyProcessed(string link)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                return context.SearchResults.Any(r => r.Link == link);
            }
        }

        private void SaveLog(string message, string level, long? crowlId = null)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                if (crowlId.HasValue && !context.Crowls.Any(c => c.Id == crowlId.Value))
                {
                    _logger.LogWarning($"CrowlId {crowlId} نامعتبر است. لاگ بدون CrowlId ذخیره می‌شود.");
                    crowlId = null;
                }

                try
                {
                    context.Logs.Add(new Log
                    {
                        Message = message,
                        Level = level,
                        CrowlId = crowlId,
                        Timestamp = DateTime.Now
                    });
                    context.SaveChanges();
                }
                catch (Exception ex)
                {
                    _logger.LogError($"خطا در ذخیره لاگ: {ex.Message}");
                }
            }
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

        private bool HandleConsentOrCaptcha(IWebDriver driver)
        {
            try
            {
                var captcha = driver.FindElement(By.CssSelector("div#recaptcha"));
                if (captcha.Displayed)
                {
                    _logger.LogInformation("کپچا شناسایی شد. نیاز به دخالت دستی.");
                    SaveLog("کپچا شناسایی شد. نیاز به دخالت دستی.", "Warning");
                    return true;
                }
            }
            catch (NoSuchElementException) { }

            try
            {
                var cookieAcceptLink = driver.FindElement(By.CssSelector("div#bnp_btn_accept a"));
                if (cookieAcceptLink.Displayed)
                {
                    cookieAcceptLink.Click();
                    Thread.Sleep(1000);
                    _logger.LogInformation("دکمه پذیرش کوکی‌ها کلیک شد.");
                    SaveLog("دکمه پذیرش کوکی‌ها کلیک شد.", "Information");
                    return false;
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
            var emailMatch = Regex.Match(html, @"([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+");
            return emailMatch.Success ? emailMatch.Value : "Not Found";
        }

        private string ExtractCategory(string html)
        {
            var result = new StringBuilder();
            if (html.Contains("فروش")) result.Append("+فروش");
            if (html.Contains("تعمیرات")) result.Append("+تعمیرات");
            if (html.Contains("نصب")) result.Append("+نصب");
            return result.Length > 0 ? result.ToString() : "نامشخص";
        }

        public IActionResult List()
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                var reports = context.Crowls.AsNoTracking().ToList();
                _logger.LogInformation($"{reports.Count} کرول از دیتابیس بازیابی شد.");
                SaveLog($"{reports.Count} کرول از دیتابیس بازیابی شد.", "Information");
                return View(reports);
            }
        }

        public IActionResult Reports(long crowlId)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                var reports = context.SearchResults
                    .Where(r => r.CrowlId == crowlId)
                    .AsNoTracking()
                    .ToList();

                var message = reports.Any()
                    ? $"{reports.Count} نتیجه برای کرول {crowlId} یافت شد."
                    : $"هیچ نتیجه‌ای برای کرول {crowlId} یافت نشد.";
                _logger.LogInformation(message);
                SaveLog(message, "Information", crowlId);

                return View(reports);
            }
        }

        public IActionResult Delete(long crowlId)
        {
            using (var scope = _scopeFactory.CreateScope())
            {
                var context = scope.ServiceProvider.GetRequiredService<SearchContext>();

                try
                {
                    var crowl = context.Crowls
                        .Include(c => c.SearchResults)
                        .Include(c => c.Logs)
                        .FirstOrDefault(c => c.Id == crowlId);

                    if (crowl == null)
                    {
                        return NotFound();
                    }

                    if (crowl.SearchResults?.Any() == true)
                    {
                        context.SearchResults.RemoveRange(crowl.SearchResults);
                    }
                    if (crowl.Logs?.Any() == true)
                    {
                        context.Logs.RemoveRange(crowl.Logs);
                    }

                    context.Crowls.Remove(crowl);
                    context.SaveChanges();
                    _logger.LogInformation($"کرول {crowlId} با موفقیت حذف شد.");
                    SaveLog($"کرول {crowlId} با موفقیت حذف شد.", "Information");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"خطا در حذف کرول {crowlId}: {ex.Message}");
                    SaveLog($"خطا در حذف کرول {crowlId}: {ex.Message}", "Error");
                    return StatusCode(500, "خطا در حذف کرول.");
                }

                return Redirect("~/home/list");
            }
        }

        [HttpGet]
        public IActionResult DownloadExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Search Results");
                var currentRow = 1;

                worksheet.Cell(currentRow, 1).Value = "عنوان";
                worksheet.Cell(currentRow, 2).Value = "خلاصه";
                worksheet.Cell(currentRow, 3).Value = "تلفن";
                worksheet.Cell(currentRow, 4).Value = "آدرس";
                worksheet.Cell(currentRow, 5).Value = "دسته‌بندی";
                worksheet.Cell(currentRow, 6).Value = "لینک";

                using (var scope = _scopeFactory.CreateScope())
                {
                    var context = scope.ServiceProvider.GetRequiredService<SearchContext>();
                    var searchResults = context.SearchResults.AsNoTracking().ToList();

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