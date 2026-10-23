namespace ExcelTransfer.Core.Workbooks.Biff12;

/// <summary>BIFF12 record type ids used by this application ([MS-XLSB] 2.3). Ids are the decoded values.</summary>
internal static class Rt
{
    public const int RowHdr = 0x0000;
    public const int CellBlank = 0x0001;
    public const int CellRk = 0x0002;
    public const int CellError = 0x0003;
    public const int CellBool = 0x0004;
    public const int CellReal = 0x0005;
    public const int CellSt = 0x0006;
    public const int CellIsst = 0x0007;
    public const int FmlaString = 0x0008;
    public const int FmlaNum = 0x0009;
    public const int FmlaBool = 0x000A;
    public const int FmlaError = 0x000B;
    public const int SSTItem = 0x0013;
    public const int FRTBegin = 0x0023;
    public const int FRTEnd = 0x0024;
    public const int ACBegin = 0x0025;
    public const int ACEnd = 0x0026;
    public const int Fmt = 0x002C;
    public const int XF = 0x002F;
    public const int ColInfo = 0x003C;
    public const int CellRString = 0x003E;
    public const int BeginSheet = 0x0081;
    public const int EndSheet = 0x0082;
    public const int BeginBook = 0x0083;
    public const int EndBook = 0x0084;
    public const int BeginBundleShs = 0x008F;
    public const int EndBundleShs = 0x0090;
    public const int BeginSheetData = 0x0091;
    public const int EndSheetData = 0x0092;
    public const int WsProp = 0x0093;
    public const int WsDim = 0x0094;
    public const int WbProp = 0x0099;
    public const int FileRecover = 0x009B;
    public const int BundleSh = 0x009C;
    public const int CalcProp = 0x009D;
    public const int BeginSst = 0x009F;
    public const int EndSst = 0x00A0;
    public const int MergeCell = 0x00B0;
    public const int BeginMergeCells = 0x00B1;
    public const int EndMergeCells = 0x00B2;
    public const int BeginList = 0x0157;
    public const int EndList = 0x0158;
    public const int BeginListCols = 0x0159;
    public const int EndListCols = 0x015A;
    public const int BeginListCol = 0x015B;
    public const int EndListCol = 0x015C;
    public const int ListCCFmla = 0x015F;
    public const int ListTrFmla = 0x0160;
    public const int BeginColInfos = 0x0186;
    public const int EndColInfos = 0x0187;
    public const int ArrFmla = 0x01AA;
    public const int ShrFmla = 0x01AB;
    public const int DataTable = 0x01AC;
    public const int SheetProtection = 0x0217;
    public const int OleSize = 0x0225;
    public const int BeginFmts = 0x0267;
    public const int EndFmts = 0x0268;
    public const int BeginCellXFs = 0x0269;
    public const int EndCellXFs = 0x026A;
    public const int BeginCellStyleXFs = 0x0272;
    public const int EndCellStyleXFs = 0x0273;
    public const int BeginListParts = 0x0294;
    public const int ListPart = 0x0295;
    public const int EndListParts = 0x0296;
    public const int SheetProtectionIso = 0x02A6;

    public static bool IsCell(int type) => (type >= CellBlank && type <= FmlaError) || type == CellRString;
    public static bool IsFormulaCell(int type) => type >= FmlaString && type <= FmlaError;
}

/// <summary>Excel error codes as stored in BIFF12 (BErr).</summary>
internal static class BErr
{
    public static string ToText(byte code) => code switch
    {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#N/A",
    };

    public static byte FromText(string text) => text switch
    {
        "#NULL!" => 0x00,
        "#DIV/0!" => 0x07,
        "#VALUE!" => 0x0F,
        "#REF!" => 0x17,
        "#NAME?" => 0x1D,
        "#NUM!" => 0x24,
        "#N/A" => 0x2A,
        "#GETTING_DATA" => 0x2B,
        _ => 0x2A,
    };
}
