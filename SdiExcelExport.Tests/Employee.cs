using SdiExcelExport.Annotations;

namespace SdiExcelExport.Tests;

public class Employee
{
    public string Name {get; set;} = string.Empty;
    public string Address {get; set;} = string.Empty;

    [OpenXmlIgnore()] /* Allow the export process to ignore */
    public bool Found { get; set; }
}