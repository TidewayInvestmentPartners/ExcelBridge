namespace ExcelBridge.Interfaces;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class EIBase : Attribute
{
    public int SheetPosition { get; set; }
    public string SheetName { get; set; }

    public string Position { get; set; }
    public string DateTimeFormat { get; set; }
}

