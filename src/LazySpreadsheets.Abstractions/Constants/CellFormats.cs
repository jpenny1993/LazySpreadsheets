namespace LazySpreadsheets.Constants;

public static class CellFormats
{
    public const string AccountingGBP = "_-\"£\"* #,##0.00_-;\\-\"£\"* #,##0.00_-;_-\"£\"* \"-\"??_-;_-@_-";
    public const string AccountingUSD = "_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* \"-\"??_ ;_-@_" ;
    public const string AccountingEUR = "_-[$?-2]\\ * #,##0.00_-;\\-[$?-2]\\ * #,##0.00_-;_-[$?-2]\\ * \"-\"??_-;_-@_-";

    public const string BooleanYN = "\"Y\";;\"N\";";
    public const string BooleanNOnly = "[=0]\"N\";";
    public const string BooleanYOnly = "[=1]\"Y\";";

    public const string BooleanYesNo = "\"YES\";;\"NO\";";
    public const string BooleanNoOnly = "[=0]\"NO\";";
    public const string BooleanYesOnly = "[=1]\"YES\";";

    public const string BooleanTickCross = "\"\u2714\ufe0f\";;\"\u2716\ufe0f\";";
    public const string BooleanTickOnly = "[=1]\"\u2714\ufe0f\";";
    public const string BooleanCrossOnly = "[=0]\"\u2716\ufe0f\";";

    public const string CurrencyGBP = "\"£\"#,##0.00";
    public const string CurrencyUSD = "[$$-409]#,##0.00";
    public const string CurrencyEUR = "[$?-2]\\ #,##0.00";

    public const string Date = "dd/mm/yyyy";
    public const string DateTime = "dd/mm/yyyy HH:mm";

    public const string LongDate = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
    public const string Time = "[$-F400]h:mm:ss\\ AM/PM";
}