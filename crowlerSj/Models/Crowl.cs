using crowlerSj.Models;
using System;
using System.Collections.Generic;
namespace crowlerSj.Models
{
    public class Crowl
{
    public long Id { get; set; }
    public string Title { get; set; }
    public DateTime InsertTime { get; set; } = DateTime.Now;
    public int CurrentPage { get; set; } = 1;
    public virtual ICollection<SearchResult> SearchResults { get; set; }
    public virtual ICollection<Log> Logs { get; set; }
}
}