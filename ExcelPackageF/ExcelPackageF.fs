namespace ExcelPackageF

open OfficeOpenXml
open System
open System.IO
open System.Xml.XPath
open System.Xml

module Excel =

    let getWorksheet (index:int) filename = 
        let file = new FileInfo(filename) 
        let xlPackage = new ExcelPackage(file)
        xlPackage.Workbook.Worksheets.[index]
        
    let getWorksheets filename = seq {
        let file = new FileInfo(filename) 
        let xlPackage = new ExcelPackage(file)
        for i in 1..xlPackage.Workbook.Worksheets.Count do
            yield xlPackage.Workbook.Worksheets.[i]
        }

    let getMaxRowNumber (worksheet:ExcelWorksheet) = 
        let nav = worksheet.WorksheetXml.CreateNavigator()
        let exp = nav.Compile("//*[name()='row']/@r")
        exp.AddSort("../@r", XmlSortOrder.Descending, XmlCaseOrder.None, "", XmlDataType.Number)
        let node = nav.SelectSingleNode(exp).UnderlyingObject :?> XmlNode
        int node.InnerText;  

    let getMaxColNumber (worksheet:ExcelWorksheet) = 
        let nav = worksheet.WorksheetXml.CreateNavigator()
        let exp = nav.Compile("//*[name()='c']/@colNumber")
        exp.AddSort("../@colNumber", XmlSortOrder.Descending, XmlCaseOrder.None, "", XmlDataType.Number)
        let node = nav.SelectSingleNode(exp).UnderlyingObject :?> XmlNode
        int node.InnerText;  

    let getContent worksheet = seq {        
        let maxRow = getMaxRowNumber worksheet
        let maxCol = getMaxColNumber worksheet
        for i in 1..maxRow do
            for j in 1..maxCol do
                let content = worksheet.Cell(i,j).Value
                yield content
    }

    let getColumn colIndex (worksheet:ExcelWorksheet) = seq { 
        let maxRow = getMaxRowNumber worksheet  
        for i in 1..maxRow do        
            let content = worksheet.Cell(i,colIndex).Value
            yield content
    }

    let getRow rowIndex (worksheet:ExcelWorksheet) = seq { 
        let maxCol = getMaxColNumber worksheet  
        for i in 1..rowIndex do        
            let content = worksheet.Cell(rowIndex,maxCol).Value
            yield content
    }