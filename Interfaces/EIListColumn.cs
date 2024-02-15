using System.Reflection;

namespace ExcelBridge.Interfaces;

public class EIListColumn
{
    public Type type { get; set; }
    public PropertyInfo pi { get; set; }
    public string ColumnName { get; set; }
    public int ColumnNumber { get; set; }
    public string DateTimeFormat { get; set; }
    public bool MustBePopulated { get; set; }
    public bool CanBeOmitted { get; set; }
    public bool Trim { get; set; }
    public bool Present { get; set; }
}