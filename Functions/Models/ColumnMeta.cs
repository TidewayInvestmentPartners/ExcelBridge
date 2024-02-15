using System.Reflection;

namespace ExcelInterface.Models;

public class ColumnMeta
{
    public Type type { get; set; }
    public PropertyInfo pi { get; set; }
    public string ColumnName { get; set; }
    public int ColumnNumber { get; set; }
    public string DateFormat { get; set; }
}

