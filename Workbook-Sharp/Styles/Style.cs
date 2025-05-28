using System.Drawing;

namespace WorkbookSharp.Styles;

public class Style : IEquatable<Style>
{
    public XlFontDecoration FontDecoration { get; set; } = XlFontDecoration.None;
    public XlFontFamily FontFamily { get; set; } = XlFontFamily.Default;
    public double? FontSize { get; set; }
    public Color FontColor { get; set; } = Color.Empty;

    public XlFontUnderlineStyle UnderlineStyle { get; set; } = XlFontUnderlineStyle.None;
    public XlFontStrikeoutStyle StrikeoutStyle { get; set; } = XlFontStrikeoutStyle.None;

    public XlDateFormat? DateFormat { get; set; }
    public string? CustomNumberFormatCode { get; set; }

    public bool UseThousandSeparator { get; set; }
    public int? DecimalPlaces { get; set; }
    public bool IsPercentage { get; set; }

    public Color FillColor { get; set; } = Color.Empty;
    public XlBorder Border { get; set; } = XlBorder.None;

    public XlHorizontalAlignment? HorizontalAlignment { get; set; }
    public XlVerticalAlignment? VerticalAlignment { get; set; }

    public bool WrapText { get; set; }
    public int? TextRotation { get; set; }
    public int? Indent { get; set; }
    public bool ShrinkToFit { get; set; }

    public Style Clone() => MemberwiseClone() as Style ?? new Style();

    public override bool Equals(object? obj) => Equals(obj as Style);

    public bool Equals(Style? other)
    {
        if (other is null) return false;

        return FontDecoration == other.FontDecoration
            && FontFamily == other.FontFamily
            && FontSize == other.FontSize
            && FontColor.ToArgb() == other.FontColor.ToArgb()
            && UnderlineStyle == other.UnderlineStyle
            && StrikeoutStyle == other.StrikeoutStyle
            && DateFormat == other.DateFormat
            && CustomNumberFormatCode == other.CustomNumberFormatCode
            && UseThousandSeparator == other.UseThousandSeparator
            && DecimalPlaces == other.DecimalPlaces
            && IsPercentage == other.IsPercentage
            && FillColor.ToArgb() == other.FillColor.ToArgb()
            && Border == other.Border
            && HorizontalAlignment == other.HorizontalAlignment
            && VerticalAlignment == other.VerticalAlignment
            && WrapText == other.WrapText
            && TextRotation == other.TextRotation
            && Indent == other.Indent
            && ShrinkToFit == other.ShrinkToFit;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(FontDecoration,
                                FontFamily,
                                FontSize,
                                FontColor.ToArgb(),
                                UnderlineStyle,
                                StrikeoutStyle,
                                DateFormat,
                                CustomNumberFormatCode)
               ^
               HashCode.Combine(UseThousandSeparator,
                                DecimalPlaces,
                                IsPercentage,
                                FillColor.ToArgb(),
                                Border,
                                HorizontalAlignment,
                                VerticalAlignment,
                                WrapText)
            ^
            HashCode.Combine(TextRotation,
                             Indent,
                             ShrinkToFit);
    }
}
