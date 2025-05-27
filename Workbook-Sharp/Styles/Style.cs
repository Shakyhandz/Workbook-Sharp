namespace WorkbookSharp.Styles;


public class Style : IEquatable<Style>
{
    public XlFontDecoration FontDecoration { get; set; } = XlFontDecoration.None;
    public XlFontFamily FontFamily { get; set; } = XlFontFamily.Default;
    public double? FontSize { get; set; }
    public System.Drawing.Color FontColor { get; set; } = System.Drawing.Color.Empty;
    public XlDateFormat? DateFormat { get; set; }
    public System.Drawing.Color FillColor { get; set; }
    public XlBorder Border { get; set; } = XlBorder.None;
    public bool UseThousandSeparator { get; set; }
    public int? decimalPlaces { get; set; }
    public bool IsPercentage { get; set; }
    public XlHorizontalAlignment? HorizontalAlignment { get; set; }
    public XlVerticalAlignment? VerticalAlignment { get; set; }

    public Style Clone() => MemberwiseClone() as Style ?? new Style();

    public override bool Equals(object? obj) => Equals(obj as Style);

    public bool Equals(Style? other)
    {
        if (other is null) return false;
        return FontDecoration == other.FontDecoration
            && FontFamily == other.FontFamily
            && FontSize == other.FontSize
            && FontColor.ToArgb() == other.FontColor.ToArgb()
            && DateFormat == other.DateFormat
            && FillColor.ToArgb() == other.FillColor.ToArgb()
            && Border == other.Border
            && UseThousandSeparator == other.UseThousandSeparator
            && decimalPlaces == other.decimalPlaces
            && IsPercentage == other.IsPercentage
            && HorizontalAlignment == other.HorizontalAlignment
            && VerticalAlignment == other.VerticalAlignment; 
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(
            FontDecoration,
            FontFamily,
            FontSize,
            FontColor.ToArgb(),
            DateFormat,
            FillColor.ToArgb(),
            Border)
            ^ 
            HashCode.Combine(            
                UseThousandSeparator,
                decimalPlaces,
                IsPercentage,
                HorizontalAlignment,
                VerticalAlignment);
    }
}