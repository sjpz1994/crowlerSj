using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace crowlerSj.Db
{
    using Microsoft.EntityFrameworkCore;

    public class SearchContext : DbContext
    {
        public SearchContext(DbContextOptions<SearchContext> options) : base(options)
        {
        }

        public DbSet<SearchResult> SearchResults { get; set; }
    }


}
