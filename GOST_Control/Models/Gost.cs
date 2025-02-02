using System;
using System.Collections.Generic;

namespace GOST_Control.Models;

public partial class Gost
{
    public int GostId { get; set; }

    public string Name { get; set; } = null!;

    public string? Description { get; set; }

    public string? FontName { get; set; }

    public double? FontSize { get; set; }

    public double? MarginTop { get; set; }

    public double? MarginBottom { get; set; }

    public double? MarginLeft { get; set; }

    public double? MarginRight { get; set; }

    public bool? PageNumbering { get; set; }

    public string? RequiredSections { get; set; }
}
