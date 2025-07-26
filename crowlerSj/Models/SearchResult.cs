using crowlerSj.Models;
using System;

public class SearchResult
{
    public long Id { get; set; }
    public string Title { get; set; }
    public string Link { get; set; }
    public string Snippet { get; set; }
    public string Phone { get; set; }
    public string Email { get; set; }
    public string Address { get; set; }
    public string Category { get; set; }
    public DateTime InsertTime { get; set; } = DateTime.Now;
    public virtual Crowl Crowl { get; set; }
    public long CrowlId { get; set; }

}
