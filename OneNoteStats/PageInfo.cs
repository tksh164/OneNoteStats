using System;

namespace OneNoteStats
{
    internal sealed class PageInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime DateTime { get; set; }
        public DateTime LastModifiedTime { get; set; }
        public int PageLevel { get; set; }
        public string IsCurrentlyViewed { get; set; }
        public string LocationPath { get; set; }
    }
}
