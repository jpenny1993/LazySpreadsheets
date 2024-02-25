namespace LazySpreadsheets.Enums;

public enum NumberFormats
{
    General = 0,                                            // General
    Integer = 1,                                            // 0
    Decimal = 2,                                            // 0.00
    IntegerWithComma = 3,                                   // #,##0
    DecimalWithComma = 4,                                   // #,##0.00
    Percentage = 9,                                         // 0%
    DecimalPercentage = 10,                                 // 0.00%
    Scientific = 11,                                        // 0.00E+00
    FractionApproximate = 12,                               // # ?/?
    FractionExact = 13,                                     // # ??/??
    Date = 14,                                              // d/m/yyyy
    DayMonthAsTextYear = 15,                                // d-mmm-yy
    DayMonthAsText = 16,                                    // d-mmm
    MonthAsTextDay = 17,                                    // mmm-yy
    Time12Hour = 18,                                        // h:mm tt
    Time12HourWithSeconds = 19,                             // h:mm:ss tt
    Time24Hour = 20,                                        // H:mm
    Time24HourWithSeconds = 21,                             // H:mm:ss
    DateTime = 22,                                          // m/d/yyyy H:mm
    CustomIntegerWithComma = 37,                            // #,##0 ;(#,##0)
    CustomIntegerWithCommaAndNegativeHightlight = 38,       // #,##0 ;[Red](#,##0)
    CustomDecimalWithComma = 39,                            // #,##0.00;(#,##0.00)
    CustomDecimalWithCommaAndNegativeHighlight = 40,        // #,##0.00;[Red](#,##0.00)
    CustomTimeMinuteSecond = 45,                            // mm:ss
    CustomTimeHourMinuteSecond = 46,                        // [h]:mm:ss
    CustomTimeMinuteSecondMillisecond = 47,                 // mmss.0
    CustomScientific = 48,                                  // ##0.0E+0
    Text = 49,                                              // @
}