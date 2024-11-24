using System;
using System.Collections.Generic;

public class Crowl
{ 
    public  long Id  { get; set; }
    public string Title { get; set; }
    public DateTime InsertTime { get; set; } = DateTime.Now;
    public virtual ICollection<SearchResult> SearchResults { get; set; }
}
