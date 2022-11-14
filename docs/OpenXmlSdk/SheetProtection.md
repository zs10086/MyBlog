# 工作表保护

使用Open Xml Sdk创建工作簿并为工作表添加保护。

<!-- tabs:start -->

#### **C#**

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace TestOpenXmlSdk;

public class Program
{
    public static void Main(string[] args)
    {
        var filepath = @"D:\TestProtection.xlsx";
        var password = "123";
        CreateSpreadsheetWorkbook(filepath, password);
    }

    /// <summary>
    /// CreateSpreadsheetWorkbook
    /// </summary>
    /// <param name="filepath"></param>
    /// <param name="password">SheetProtection Cleartext password.</param>
    static void CreateSpreadsheetWorkbook(string filepath, string password)
    {
        // Create a spreadsheet document by supplying the filepath.
        // By default, AutoSave = true, Editable = true, and Type = xlsx.
        SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
            Create(filepath, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add a WorksheetPart to the WorkbookPart.
        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
            AppendChild(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet()
        {
            Id = spreadsheetDocument.WorkbookPart.
            GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "mySheet"
        };
        sheets.Append(sheet);

        // Add SheetProtection to the worksheet.
        var sp = worksheetPart.Worksheet.Descendants<SheetProtection>().FirstOrDefault();
        if (sp == null)
        {
            SheetProtection sheetProtection = new SheetProtection();
            sheetProtection.Sheet = true;
            sheetProtection.Scenarios = true;
            sheetProtection.Objects = true;
            sheetProtection.Password = GetPasswordHash(password);

            worksheetPart.Worksheet.InsertAfter(sheetProtection, worksheetPart.Worksheet.Descendants<SheetData>().LastOrDefault());
        }

        workbookpart.Workbook.Save();

        // Close the document.
        spreadsheetDocument.Close();
    }

    /// <summary>
    /// ECMA-376 hash algorithm. Read ECMA-376 Part 4 <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376"/> for more infomation.
    /// </summary>
    /// <param name="password">Cleartext password.</param>
    /// <returns>Hashed password.</returns>
    static string GetPasswordHash(string password)
    {
        byte[] passwordCharacters = Encoding.ASCII.GetBytes(password);
        int hash = 0;
        if (passwordCharacters.Length > 0)
        {
            int charIndex = passwordCharacters.Length;

            while (charIndex-- > 0)
            {
                hash = hash >> 14 & 0x01 | hash << 1 & 0x7fff;
                hash ^= passwordCharacters[charIndex];
            }
            hash = hash >> 14 & 0x01 | hash << 1 & 0x7fff;
            hash ^= passwordCharacters.Length;
            hash ^= 0x8000 | 'N' << 8 | 'K';
        }
        return Convert.ToString(hash, 16).ToUpperInvariant();
    }
}
```

#### **VB**

```vb
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.Text
 
Dim TestOpenXmlSdk As Namespace
Dim TestOpenXmlSdk As namespace
End Namespace
 
Public Class Program
    Public Shared  Sub Main(ByVal args() As String)
        Dim filepath As var =  "D:\TestProtection.xlsx" 
        Dim password As var =  "123" 
        CreateSpreadsheetWorkbook(filepath, password)
    End Sub
 
    '/ <summary>
    '/ CreateSpreadsheetWorkbook
    '/ </summary>
    '/ <param name="filepath"></param>
    '/ <param name="password">SheetProtection Cleartext password.</param>
    Shared  Sub CreateSpreadsheetWorkbook(ByVal filepath As String, ByVal password As String)
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
            Create(filepath, SpreadsheetDocumentType.Workbook)
 
        ' Add a WorkbookPart to the document.
        Dim workbookpart As WorkbookPart =  spreadsheetDocument.AddWorkbookPart() 
        workbookpart.Workbook = New Workbook()
 
        ' Add a WorksheetPart to the WorkbookPart.
        Dim worksheetPart As WorksheetPart =  workbookpart.AddNewPart<WorksheetPart>() 
        worksheetPart.Worksheet = New Worksheet(New SheetData())
 
        ' Add Sheets to the Workbook.
        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
            AppendChild(New Sheets())
 
        ' Append a new worksheet and associate it with the workbook.
        Sheet sheet = Function Sheet() As Shadows
            Id = spreadsheetDocument.WorkbookPart.
            GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "mySheet"
        End Function

        sheets.Append(sheet)
 
        ' Add SheetProtection to the worksheet.
        Dim sp As var =  worksheetPart.Worksheet.Descendants<SheetProtection>().FirstOrDefault() 
        If sp Is Nothing Then
            Dim sheetProtection As SheetProtection =  New SheetProtection() 
            sheetProtection.Sheet = True
            sheetProtection.Scenarios = True
            sheetProtection.Objects = True
            sheetProtection.Password = GetPasswordHash(password)
 
            worksheetPart.Worksheet.InsertAfter(sheetProtection, worksheetPart.Worksheet.Descendants<SheetData>().LastOrDefault())
        End If
 
        workbookpart.Workbook.Save()
 
        ' Close the document.
        spreadsheetDocument.Close()
    End Sub
 
    '/ <summary>
    '/ ECMA-376 hash algorithm. Read ECMA-376 Part 4 <see href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376"/> for more infomation.
    '/ </summary>
    '/ <param name="password">Cleartext password.</param>
    '/ <returns>Hashed password.</returns>
    Shared Function GetPasswordHash(ByVal password As String) As String
        Dim passwordCharacters() As Byte =  Encoding.ASCII.GetBytes(password) 
        Dim hash As Integer =  0 
        If passwordCharacters.Length > 0 Then
            Dim charIndex As Integer =  passwordCharacters.Length 
 
            While charIndex = While charIndex - 1
                hash = hash > 14 & 0x01 | hash < 1 & 0x7fff
                Dim ^ As hash =  passwordCharacters(charIndex) 
            End While
            hash = hash > 14 & 0x01 | hash < 1 & 0x7fff
            Dim ^ As hash =  passwordCharacters.Length 
            Dim ^ As hash =  0x8000 | "N"c < 8 | "K"c 
        End If
        Return Convert.ToString(hash,16).ToUpperInvariant()
    End Function
End Class

'----------------------------------------------------------------
' Converted from C# to VB .NET using CSharpToVBConverter(1.2).
' Developed by: Kamal Patel (http://www.KamalPatel.net) 
'----------------------------------------------------------------

```

<!-- tabs:end -->

![logo](../_images/GetPasswordHash.png ':size=50%')
