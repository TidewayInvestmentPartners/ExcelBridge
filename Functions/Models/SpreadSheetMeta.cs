namespace ExcelBridge.Models;

public class SpreadSheetMeta
{
    public string SheetName { get; set; }
    public int SheetPosition { get; set; }
    public int HeaderRow { get; set; }
    public List<ColumnMeta> Columns { get; } = new List<ColumnMeta>();
}

