namespace ExcelBridge.Interfaces;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class EIList : Attribute
{
    public int SheetPosition { get; set; }
    public string SheetName { get; set; }

    public int HeaderRow { get; set; }
    public int StartColumn { get; set; }
}

