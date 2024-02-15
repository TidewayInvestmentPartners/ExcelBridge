namespace ExcelBridge.Interfaces;

public class EIColumn : Attribute
{
    public string Header { get; set; }
    //public bool ColonAtTheEnd { get; set; } = false;
    public bool CanBeOmitted { get; set; }
    public string DateTimeFormat { get; set; }
    public bool MustBePopulated { get; set; }
    public bool Trim { get; set; } = false;

}

