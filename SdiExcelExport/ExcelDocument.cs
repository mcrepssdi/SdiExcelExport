using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SdiExcelExport.Annotations;

namespace SdiExcelExport;

public static class ExcelDocument
{
    private enum ExcelStyleNames
    {
        Default,
        ColHeaders,
        Gray,
        LightSteelBlue,
    }
    
    /// <summary>
    /// 
    /// </summary>
    /// <param name="items"></param>
    /// <param name="fileName"></param>
    /// <param name="alternateColors"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    public static (string fileName, MemoryStream documentStream) GenerateExcelDoc<T>(this IEnumerable<T> items, string fileName, bool alternateColors = false)
    {
        MemoryStream documentStream = new ();
        SpreadsheetDocument document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylePart.Stylesheet = new ExcelExportStyles().GenerateStyleSheet();
        stylePart.Stylesheet.Save();
        
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        SheetData sheetData = new ();
        worksheetPart.Worksheet = new Worksheet(sheetData);
 
        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
        Sheet sheet = new () { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = fileName };
        sheets.Append(sheet);
        
        Row headerRow = new();
        Cell cell;
        PropertyInfo[] props = typeof(T).GetProperties();
        foreach (PropertyInfo prop in props)
        {
            List<Attribute> attrs = Attribute.GetCustomAttributes(prop).ToList();
            Attribute? ignoreProperty = (from s in attrs where s is OpenXmlIgnore select s).FirstOrDefault();
            if (ignoreProperty != null) continue;
            
            cell = ConstructCell(prop.Name, CellValues.String, ExcelStyleNames.ColHeaders);
            headerRow.AppendChild(cell);
        }
        sheetData.AppendChild(headerRow);

        int rowNum = 1;
        foreach (T item in items)
        {
            if (item == null) continue;
            
            Row newRow = new ();
            ExcelStyleNames cellStyle = ExcelStyleNames.Default;
            if (alternateColors)
            {
                cellStyle = ExcelStyleNames.Default;
                if (rowNum % 2 == 0)
                {
                    cellStyle = ExcelStyleNames.LightSteelBlue;
                }
            }

            foreach (PropertyInfo prop in props)
            {
                List<Attribute> attrs = Attribute.GetCustomAttributes(prop).ToList();
                Attribute? ignoreProperty = (from s in attrs where s is OpenXmlIgnore select s).FirstOrDefault();
                if (ignoreProperty != null) continue;
                
                Type? nullable = Nullable.GetUnderlyingType(prop.PropertyType);
                cell = InitCell(nullable ?? prop.PropertyType, prop, cellStyle, item);
                newRow.AppendChild(cell);
            }
            sheetData.AppendChild(newRow);
            rowNum++;
        }
        document.Close();/* Closes stream and write data to the buffer NOTE: Important do not use Save on workbook */
        return ($"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx", documentStream);
    }

    private static Cell ConstructCell<T>(string value, CellValues dataType, T styleIndex) where T : Enum
    {
        return new Cell
        {
            CellValue = new CellValue(value),
            DataType = new EnumValue<CellValues>(dataType),
            StyleIndex = Convert.ToUInt32(EnumToInt(styleIndex))
        };
    }

    private static Cell ConstructCell<T>(double value, CellValues dataType,  T styleIndex) where T : Enum
    {
        return new Cell
        {
            CellValue = new CellValue(value),
            DataType = new EnumValue<CellValues>(dataType),
            StyleIndex = Convert.ToUInt32(EnumToInt(styleIndex))
        };
    }

    private static Cell ConstructCell<T>(DateTime value, CellValues dataType,  T styleIndex) where T : Enum
    {
        return new Cell
        {
            CellValue = new CellValue(value),
            DataType = new EnumValue<CellValues>(dataType),
            StyleIndex = Convert.ToUInt32(EnumToInt(styleIndex))
        };
    }
    
    private static Cell InitCell(Type inType, PropertyInfo prop, ExcelStyleNames styleName, object item)
    {
        Cell cell = new ();
        Type? type = GetDataType(inType);
        if (type == typeof(string))
        {
             object value = prop.GetValue(item) ?? string.Empty;
             cell = ConstructCell((string)value, CellValues.String, styleName);
        }
        else if (type == typeof(int))
        {
            object value = prop.GetValue(item) ?? 0;
            cell = ConstructCell((int)value, CellValues.Number, styleName);
        }
        else if (type == typeof(long))
        {
            object value = prop.GetValue(item) ?? 0;
            cell = ConstructCell((long) value, CellValues.Number, styleName);
        }
        else if (type == typeof(DateTime))
        {
            object value = prop.GetValue(item) ?? new DateTime(1900,1,1);
            DateTime dt = (DateTime)value;
            cell = ConstructCell(dt.ToString("MM/dd/yyyy"), CellValues.String, styleName);
        }
        else if (type == typeof(decimal))
        {
            object value = prop.GetValue(item) ?? 0;
            cell = ConstructCell((double) value, CellValues.Number, styleName);    
        }
        else if (type == typeof(double))
        {
            object value = prop.GetValue(item) ?? 0;
            cell = ConstructCell((double) value, CellValues.Number, styleName);         
        }
        else if (type == typeof(bool))
        {
            object value = prop.GetValue(item) ?? string.Empty;
            cell = ConstructCell((string)value, CellValues.String, styleName);
        }
        return cell;
    }
    
    private static int EnumToInt<TValue>(TValue value) where TValue : Enum => (int)(object)value;

    private static Type? GetDataType(Type type)
    {
        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            Type? nullableType =  Nullable.GetUnderlyingType(type);
            return nullableType;
        }

        return type;
    }
}