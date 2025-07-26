using crowlerSj.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace crowlerSj.Db
{
    public class SearchContext : DbContext
    {
        public SearchContext(DbContextOptions<SearchContext> options) : base(options)
        {
        }

        public DbSet<SearchResult> SearchResults { get; set; }
        public DbSet<Crowl> Crowls { get; set; }
        public DbSet<Setting> Settings { get; set; }
        public DbSet<Log> Logs { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // Seed Data
            SeedData(modelBuilder);

            // تنظیم رابطه Logs و Crowls
            modelBuilder.Entity<Log>()
                .HasOne(l => l.Crowl)
                .WithMany(c => c.Logs)
                .HasForeignKey(l => l.CrowlId)
                .IsRequired(false); // CrowlId اختیاریه

            // تنظیم رابطه SearchResults و Crowls
            modelBuilder.Entity<SearchResult>()
                .HasOne(r => r.Crowl)
                .WithMany(c => c.SearchResults)
                .HasForeignKey(r => r.CrowlId)
                .IsRequired(true); // CrowlId برای SearchResult اجباریه
        }

        private void SeedData(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Setting>().HasData(new Setting
            {
                Id = 1,
                IsCrowl = false
            });
        }
    }
}