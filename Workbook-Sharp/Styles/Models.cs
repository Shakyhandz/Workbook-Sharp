namespace WorkbookSharp.Styles;

[Flags]
public enum XlBorder
{
    None = 0,
    Top = 1,
    Bottom = 2,
    Left = 4,
    Right = 8,
    Around = Top | Bottom | Left | Right,
}

[Flags]
public enum XlFontDecoration
{
    None = 0,
    Bold = 1,
    Italic = 2,
    Underline = 4,
    Strikeout = 8,
}

// TODO: translate to EnumValue<UnderlineValues>
public enum XlFontUnderlineStyle
{
    None = 0,
    Single = 1,
    Double = 2,
    SingleAccounting = 3,
    DoubleAccounting = 4,
}

public enum XlHorizontalAlignment
{
    Left = 0,
    Right = 1,
    Center = 2,
}

public enum XlFontStrikeoutStyle
{
    None = 0,
    Single = 1,
    Double = 2,
}

public enum XlFontFamily
{
    Default = 0,
    Calibri = 1,
    Arial = 2,
}

public enum XlDateFormat
{
    Date = 0,
    //Hours = 1,
    HoursMinutes = 2,
    HoursMinutesSeconds = 3,
    DateHours = 4,
    DateHoursMinutes = 5,
    DateHoursMinutesSeconds = 6,
    None = 7,
}
