﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookSharp.Styles;

internal class StyleManager
{
    private readonly Dictionary<Style, uint> _styleIndexMap = new();
    private readonly Dictionary<string, uint> _fontMap = new();
    private readonly Dictionary<string, uint> _fillMap = new();
    private readonly Dictionary<string, uint> _borderMap = new();
    private readonly List<Font> _fonts = new();
    private readonly List<Fill> _fills = new();
    private readonly List<Border> _borders = new();
    private readonly List<CellFormat> _cellFormats = new();
    private readonly List<NumberingFormat> _numberFormats = new();
    
    private readonly HashSet<uint> _numberStyleIndexes = new();
    private readonly HashSet<uint> _textDecorationIndexes = new();
    internal HashSet<uint> GetNumberStyles() => _numberStyleIndexes.ToHashSet();
    internal HashSet<uint> GetTextDecorations() => _textDecorationIndexes.ToHashSet();

    private uint _nextNumberFormatId = 164; // Built-in formats end at 163

    public StyleManager()
    {
        // Just a stupid hardcoded Excel thing that these are the first two fills even if you override them

        // Fill index 0
        var fill0 = new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } };
        _fillMap[fill0.OuterXml] = 0;
        _fills.Add(fill0);

        // Fill index 1
        var fill1 = new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } };
        _fillMap[fill1.OuterXml] = 1;
        _fills.Add(fill1);

        // Add default empty style (index 0)
        _ = GetStyleIndex(new Style());
    }

    internal uint GetStyleIndex(Style style)
    {
        // Clone to avoid modifying the key style
        style = style.Clone();

        if (_styleIndexMap.TryGetValue(style, out var index))
            return index;

        var fontId = AddFont(style);
        var fillId = AddFill(style);
        var borderId = AddBorder(style);
        var numberFormatId = GetNumberFormatId(style);

        var cellFormat = new CellFormat
        {
            FontId = fontId,
            FillId = fillId,
            ApplyFill = style.FillColor != default,
            BorderId = borderId,
            ApplyBorder = style.Border != XlBorder.None,
            NumberFormatId = numberFormatId,
            ApplyNumberFormat = numberFormatId >= 164,         
        };

        var alignment = style.HorizontalAlignment switch
        {
            XlHorizontalAlignment.Center => new Alignment { Horizontal = HorizontalAlignmentValues.Center },
            XlHorizontalAlignment.Right => new Alignment { Horizontal = HorizontalAlignmentValues.Right },
            XlHorizontalAlignment.Left => new Alignment { Horizontal = HorizontalAlignmentValues.Left },
            _ => null
        };

        if (alignment != null)
        {
            cellFormat.ApplyAlignment = true;
            cellFormat.Alignment = alignment;
        }

        _cellFormats.Add(cellFormat);
        index = (uint)_cellFormats.Count - 1;
        _styleIndexMap[style] = index;

        return index;
    }

    internal Stylesheet BuildStylesheet()
    {
        var stylesheet = new Stylesheet
        {
            Fonts = new Fonts(_fonts) { Count = (uint)_fonts.Count },
            Fills = new Fills(_fills) { Count = (uint)_fills.Count },
            Borders = new Borders(_borders) { Count = (uint)_borders.Count },
            CellFormats = new CellFormats(_cellFormats) { Count = (uint)_cellFormats.Count },
        };

        if (_numberFormats.Any())
        {
            stylesheet.NumberingFormats = new NumberingFormats(_numberFormats)
            {
                Count = (uint)_numberFormats.Count
            };
        }

        return stylesheet;
    }

    private uint AddFont(Style style)
    {
        var font = new Font();

        // TODO: font color
        //font.Append(new Color { Rgb = new HexBinaryValue { Value = "FF0000" } }); // Red

        if (style.FontSize is double size)
            font.Append(new FontSize { Val = size });

        if (style.FontFamily != XlFontFamily.Default)
            font.Append(new FontName { Val = style.FontFamily.ToString() });

        if (style.FontDecoration.HasFlag(XlFontDecoration.Bold)) 
            font.Append(new Bold());

        if (style.FontDecoration.HasFlag(XlFontDecoration.Italic)) 
            font.Append(new Italic());

        if (style.FontDecoration.HasFlag(XlFontDecoration.Underline)) 
            font.Append(new Underline());

        if (style.FontDecoration.HasFlag(XlFontDecoration.Strikeout)) 
            font.Append(new Strike());

        var key = font.OuterXml;
        if (_fontMap.TryGetValue(key, out var existingId))
            return existingId;

        _fonts.Add(font);
        var id = (uint)_fonts.Count - 1;
        _fontMap[key] = id;
        return id;
    }

    private uint AddFill(Style style)
    {
        Fill fill;

        if (style.FillColor == default)
        {
            fill = new Fill(new PatternFill { PatternType = PatternValues.None });
        }
        else
        {
            var foregroundColor = TranslateForeground(style.FillColor);

            fill = new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = foregroundColor,
                    BackgroundColor = new BackgroundColor { Rgb = foregroundColor.Rgb }
                }
            };
        }

        var key = fill.OuterXml;
        if (_fillMap.TryGetValue(key, out var existingId))
            return existingId;

        _fills.Add(fill);
        var id = (uint)_fills.Count - 1;
        _fillMap[key] = id;
        return id;
    }

    private static Dictionary<System.Drawing.Color, string> _foregroundCache = new();

    private ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
    {
        if (_foregroundCache.TryGetValue(fillColor, out var cachedValue))
            return new ForegroundColor { Rgb = new HexBinaryValue { Value = cachedValue } };

        var color = System.Drawing
                          .ColorTranslator
                          .ToHtml(System.Drawing.Color.FromArgb(fillColor.A, fillColor.R, fillColor.G, fillColor.B))
                          .Replace("#", "");

        _foregroundCache[fillColor] = color;

        return new ForegroundColor { Rgb = new HexBinaryValue { Value = color } };
    }
    

    private uint AddBorder(Style style)
    {
        var border = new Border
        {
            TopBorder = new TopBorder { Style = style.Border.HasFlag(XlBorder.Top) ? BorderStyleValues.Thin : BorderStyleValues.None },
            BottomBorder = new BottomBorder { Style = style.Border.HasFlag(XlBorder.Bottom) ? BorderStyleValues.Thin : BorderStyleValues.None },
            LeftBorder = new LeftBorder { Style = style.Border.HasFlag(XlBorder.Left) ? BorderStyleValues.Thin : BorderStyleValues.None },
            RightBorder = new RightBorder { Style = style.Border.HasFlag(XlBorder.Right) ? BorderStyleValues.Thin : BorderStyleValues.None },
            DiagonalBorder = new DiagonalBorder()
        };

        var key = border.OuterXml;
        if (_borderMap.TryGetValue(key, out var existingId))
            return existingId;

        _borders.Add(border);
        var id = (uint)_borders.Count - 1;
        _borderMap[key] = id;
        return id;
    }

    /**********************************
        https://jason-ge.medium.com/create-excel-using-openxml-in-net-6-3b601ddf48f7
         
        ID  FORMAT CODE
        0   General
        1   0
        2   0.00
        3   #,##0
        4   #,##0.00
        9   0%
        10  0.00%
        11  0.00E+00
        12  # ?/?
        13  # ??/??
        14  d/m/yyyy
        15  d-mmm-yy
        16  d-mmm
        17  mmm-yy
        18  h:mm tt
        19  h:mm:ss tt
        20  H:mm
        21  H:mm:ss
        22  m/d/yyyy H:mm
        37  #,##0 ;(#,##0)
        38  #,##0 ;[Red](#,##0)
        39  #,##0.00;(#,##0.00)
        40  #,##0.00;[Red](#,##0.00)
        45  mm:ss
        46  [h]:mm:ss
        47  mmss.0
        48  ##0.0E+0
        49  @
    ***********************************/

    private uint GetNumberFormatId(Style style)
    {
        if (style.IsPercentage)
        {
            if (style.decimalPlaces == null || style.decimalPlaces == 0)
                return 9;

            if (style.decimalPlaces == 2)
                return 10;


            var id = _nextNumberFormatId++;

            _numberFormats.Add(new NumberingFormat
            {
                NumberFormatId = id,
                FormatCode = "0." + new string('0', style.decimalPlaces ?? 2) + "%"
            });

            return id;
        }

        if (style.UseThousandSeparator)
        {
            if (style.decimalPlaces == null || style.decimalPlaces == 0)
                return 3;

            if (style.decimalPlaces == 2)
                return 4;

            var id = _nextNumberFormatId++;

            _numberFormats.Add(new NumberingFormat
            {
                NumberFormatId = id,
                FormatCode = "#,##0" + (style.decimalPlaces.HasValue ? "." + new string('0', style.decimalPlaces.Value) : "")
            });

            return id;
        }

        if (style.decimalPlaces != null && style.decimalPlaces > 0)
        {
            var id = _nextNumberFormatId++;
            var formatCode = "#,##0.0";

            if (style.decimalPlaces > 1)            
                formatCode += new string('#', style.decimalPlaces.Value - 1);
            
            _numberFormats.Add(new NumberingFormat
            {
                NumberFormatId = id,
                FormatCode = formatCode,
            });

            return id;
        }

        return style.DateFormat switch
        {
            XlDateFormat.Date => 14,
            //XlDateFormat.Hours => ,
            XlDateFormat.HoursMinutes => 20,
            XlDateFormat.HoursMinutesSeconds => 21,
            XlDateFormat.DateHours => AddCustomNumberFormat("yyyy-mm-dd hh"),
            XlDateFormat.DateHoursMinutes => 22,
            XlDateFormat.DateHoursMinutesSeconds => AddCustomNumberFormat("yyyy-mm-dd hh:mm:ss"),
            _ => 0
        };
    }
        
    private uint AddCustomNumberFormat(string formatCode)
    {
        var id = _nextNumberFormatId++;

        _numberFormats.Add(new NumberingFormat
        {
            NumberFormatId = id,
            FormatCode = formatCode,
        });

        return id;
    }
     
}