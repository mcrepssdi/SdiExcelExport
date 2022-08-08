# OpenXMl Library 
A library used to create OpenXml documents


# Usage
```c#
        public Employee{
            public string Name {get;set;} = string.Empty;
            public string Address {get;set;} = string.Empty;
            
            [OpenXmlIgnore()] /* Allow the export process to ignore */
            public bool found {get;set;}
        }

        List<Employee> list = new()
        {
            new Employee {Name = "Merl", Address = "Somewhere"},
            new Employee {Name = "Eric", Address = "Somewhere"},
            new Employee {Name = "Chad", Address = "Somewhere"}
        };
        
        (string fileName, MemoryStream documentStream) 
            = list.GenerateExcelDoc( "ExcelFileName");
        
        FileStream outStream = File.OpenWrite($"C:\\Users\\mcreps\\Documents\\{fileName}");
        documentStream.WriteTo(outStream);
        outStream.Flush();
        outStream.Close();
```

# Dependencies
1. DocumentFormat.OpenXml 2.17.1