using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using WorkbookSharp.Styles;

namespace WorkbookSharp.Cells;

internal class CellRichText : CellAction
{
    private InlineString? _inlineString;
    internal InlineString? InlineString => _inlineString;

    internal List<(string text, Style? style)> Runs = [];

    internal CellRichText(string address, params (string text, Style style)[] runs) : base(address, null) => SetInlineString(runs);

    internal CellRichText((uint row, uint column) cellReference, params (string text, Style style)[] runs) : base(cellReference, null) => SetInlineString(runs);

    public void SetInlineString(params (string text, Style style)[] runs)
    {
        Runs = runs.Select(x => (x.text, (Style?)x.style)).ToList();

        var inlineString = new InlineString();

        foreach (var (text, style) in runs)
        {
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var rPr = new RunProperties();

            if (style.FontSize != null)
                rPr.Append(new FontSize { Val = style.FontSize });

            if (style.FontColor != System.Drawing.Color.Empty)
            {
                var colorHex = $"{style.FontColor.R:X2}{style.FontColor.G:X2}{style.FontColor.B:X2}";
                rPr.Append(new Color { Rgb = colorHex });
            }

            if (style.FontDecoration.HasFlag(XlFontDecoration.Bold))
                rPr.Append(new Bold());

            if (style.FontDecoration.HasFlag(XlFontDecoration.Italic))
                rPr.Append(new Italic());

            run.PrependChild(rPr);
            inlineString.Append(run);
        }

        _inlineString = inlineString;
    }
}
