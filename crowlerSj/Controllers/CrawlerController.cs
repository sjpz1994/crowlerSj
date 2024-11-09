using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System;

public class CrawlerController : Controller
{
    private static readonly HttpClient client = new HttpClient();

    public async Task<IActionResult> Index()
    {
        string query = "دوربین مدار بسته خرید فروش تعمیرات نصب";
        string url = $"https://www.google.com/search?q={query}";

        var html = await GetHtmlAsync(url);
        var results = ExtractSearchResults(html);

        foreach (var result in results)
        {
            var detailHtml = await GetHtmlAsync(result.Link);
            ExtractDetails(detailHtml, result);
        }

        return View(results);
    }

    private async Task<string> GetHtmlAsync(string url)
    {
        client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
        var response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();
        await Task.Delay(2000); // تاخیر برای جلوگیری از مسدود شدن
        return await response.Content.ReadAsStringAsync();
    }

    private List<SearchResult> ExtractSearchResults(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var results = new List<SearchResult>();
        var nodes = doc.DocumentNode.SelectNodes("//div[@class='tF2Cxc']");

        if (nodes != null)
        {
            foreach (var node in nodes)
            {
                var titleNode = node.SelectSingleNode(".//h3");
                var linkNode = node.SelectSingleNode(".//a");
                var snippetNode = node.SelectSingleNode(".//span[@class='aCOpRe']");

                if (titleNode != null && linkNode != null && snippetNode != null)
                {
                    results.Add(new SearchResult
                    {
                        Title = titleNode.InnerText,
                        Link = linkNode.GetAttributeValue("href", ""),
                        Snippet = snippetNode.InnerText
                    });
                }
            }
        }

        return results;
    }

    private void ExtractDetails(string html, SearchResult result)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        // استخراج شماره تماس و آدرس
        var phoneMatch = Regex.Match(html, @"\(?\d{3}\)?-?\s*\d{3}-\d{4}");
        var addressNode = doc.DocumentNode.SelectSingleNode("//address");

        result.Phone = phoneMatch.Success ? phoneMatch.Value : "Not Found";
        result.Address = addressNode != null ? addressNode.InnerText.Trim() : "Not Found";

        // دسته‌بندی (فرض کنید این اطلاعات در متن صفحه موجود است)
        if (html.Contains("فروش"))
        {
            result.Category = "فروش";
        }
        else if (html.Contains("تعمیرات"))
        {
            result.Category = "تعمیرات";
        }
        else if (html.Contains("نصب"))
        {
            result.Category = "نصب";
        }
        else
        {
            result.Category = "نامشخص";
        }
    }
}

public class SearchResult
{
    public int Id { get; set; }
    public string Title { get; set; }
    public string Link { get; set; }
    public string Snippet { get; set; }
    public string Phone { get; set; }
    public string Email { get; set; }
    public string Address { get; set; }
    public string Category { get; set; }
    public DateTime InsertTime { get; set; } = DateTime.Now;
}
