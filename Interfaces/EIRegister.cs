namespace ExcelBridge.Interfaces;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class EIRegister : Attribute
{
    public int SheetPosition { get; set; }
    public string SheetName { get; set; }

    public string StartPosition { get; set; }
    public string EndPosition { get; set; }
    public int ColumnSeparation { get; set; }
}