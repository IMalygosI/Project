using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GOST_Control
{
    public struct PageSizesRazmer2
    {
        public string Name { get; set; }
        public double Width { get; set; } 
        public double Height { get; set; } 
    }

    public static class PageSizesRazmer
    {
        public static readonly PageSizesRazmer2[] AllSizes = new PageSizesRazmer2[]
        {
        new PageSizesRazmer2 { Name = "US Letter", Width = 21.59, Height = 27.94 },
        new PageSizesRazmer2 { Name = "US Legal", Width = 21.59, Height = 35.56 },
        new PageSizesRazmer2 { Name = "A4", Width = 21, Height = 29.7 },
        new PageSizesRazmer2 { Name = "A5", Width = 14.8, Height = 21 },
        new PageSizesRazmer2 { Name = "B5", Width = 17.6, Height = 25 },
        new PageSizesRazmer2 { Name = "Envelope #10", Width = 10.48, Height = 24.13 },
        new PageSizesRazmer2 { Name = "Envelope DL", Width = 11, Height = 22 },
        new PageSizesRazmer2 { Name = "Tabloid", Width = 27.94, Height = 43.18 },
        new PageSizesRazmer2 { Name = "A3", Width = 29.7, Height = 42 },
        new PageSizesRazmer2 { Name = "Tabloid Oversize", Width = 30.48, Height = 45.71 },
        new PageSizesRazmer2 { Name = "ROC 16K", Width = 19.68, Height = 27.3 },
        new PageSizesRazmer2 { Name = "Envelope Choukei 3", Width = 11.99, Height = 23.49 },
        new PageSizesRazmer2 { Name = "Super B/A3", Width = 33.02, Height = 48.25 }
        };
    }
}
