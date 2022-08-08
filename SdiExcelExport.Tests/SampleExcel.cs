using NUnit.Framework;

namespace SdiExcelExport.Tests;

public class SampleExcel
{
    [SetUp]
    public void Setup(){}

    [TearDown]
    public void Teardown(){}


    [Test]
    public void MakeDoc_NoStrpping_ExpectTrue()
    {
        List<Employee> list = new()
        {
            new Employee {Name = "Merl", Address = "Somewhere"},
            new Employee {Name = "Eric", Address = "Somewhere"},
            new Employee {Name = "Chad", Address = "Somewhere"}
        };
        
        (string fileName, MemoryStream documentStream) 
            = list.GenerateExcelDoc( "ExcelFileName");
        Assert.IsNotEmpty(fileName);
        
        FileStream outStream = File.OpenWrite($"C:\\Users\\mcreps\\Documents\\{fileName}");
        documentStream.WriteTo(outStream);
        outStream.Flush();
        outStream.Close();
        Assert.Pass("Passed");
    }
    
    
    [Test]
    public void MakeDoc_WithStripping_ExpectTrue()
    {
        List<Employee> list = new()
        {
            new Employee {Name = "Merl", Address = "Somewhere"},
            new Employee {Name = "Eric", Address = "Somewhere"},
            new Employee {Name = "Chad", Address = "Somewhere"}
        };
        
        (string fileName, MemoryStream documentStream) 
            = list.GenerateExcelDoc( "ExcelFileNameWithStripping", true);
        Assert.IsNotEmpty(fileName);
        
        FileStream outStream = File.OpenWrite($"C:\\Users\\mcreps\\Documents\\{fileName}");
        documentStream.WriteTo(outStream);
        outStream.Flush();
        outStream.Close();
        Assert.Pass("Passed");
    }
}