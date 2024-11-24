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
        public DbSet<Crowl> Crowls { get; set; }
        public DbSet<Setting> Settings { get; set; }
   
    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        //seed Data
        SeedData(modelBuilder);
    }
    private void SeedData(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Setting>().HasData(new Setting
        {
            Id = 1,
            IsCrowl=false
        });
    }
}
}