namespace ExcelPackageF

open OfficeOpenXml
open System
open System.IO
open System.Xml.XPath
open System.Xml

module Excel =

    /// <summary>Reads an Excel file and returns the worksheet with the specified index</summary>
    /// <param name="index">The index of the worksheet in the file (starting from 1!)</param>
    /// <param name="filename">The input file.</param>
    /// <returns>Excel worksheet</returns>
    let getWorksheetByIndex (index:int) filename = 
        let file = new FileInfo(filename) 
        let xlPackage = new ExcelPackage(file)
        xlPackage.Workbook.Worksheets.[index]

    /// <summary>Reads an Excel file and returns the worksheet with the specified name</summary>
    /// <param name="name">The name of the worksheet in the file</param>
    /// <param name="filename">The input file.</param>
    /// <returns>Excel worksheet</returns>
    let getWorksheetByName (name:string) filename = 
        let file = new FileInfo(filename) 
        let xlPackage = new ExcelPackage(file)
        xlPackage.Workbook.Worksheets.[name]

    /// <summary>Reads an Excel file and returns a sequence of all the worksheets in the file</summary>    
    /// <param name="filename">The input file.</param>
    /// <returns>Sequence of Excel worksheets</returns>        
    let getWorksheets filename = seq {
        let file = new FileInfo(filename) 
        let xlPackage = new ExcelPackage(file)
        for i in 1..xlPackage.Workbook.Worksheets.Count do
            yield xlPackage.Workbook.Worksheets.[i]
        }

    /// <summary>Gets the maximum row number for a given worksheet</summary>    
    /// <param name="worksheet">The input worksheet.</param>
    /// <returns>Maximum row number</returns>
    let getMaxRowNumber (worksheet:ExcelWorksheet) = 
        worksheet.Dimension.End.Row 

    /// <summary>Gets the maximum column number for a given worksheet</summary>    
    /// <param name="worksheet">The input worksheet.</param>
    /// <returns>Maximum column number</returns>
    let getMaxColNumber (worksheet:ExcelWorksheet) = 
        worksheet.Dimension.End.Column

    /// <summary>Gets all the values from all the cells in a given worksheet in a sequence. The traversal is done line by line</summary>    
    /// <param name="worksheet">The input worksheet.</param>
    /// <returns>Sequence of cell values</returns>
    let getContent worksheet = seq {        
        let maxRow = getMaxRowNumber worksheet
        let maxCol = getMaxColNumber worksheet
        for i in 1..maxRow do
            for j in 1..maxCol do
                let content = worksheet.Cells.[i,j].Value
                yield content
    }

    /// <summary>Gets all the values from given column in a given worksheet in a sequence. </summary>   
    /// <param name="colIndex">The column index (starting from 1!).</param> 
    /// <param name="worksheet">The input worksheet.</param>
    /// <returns>Sequence of cell values</returns>
    let getColumn colIndex (worksheet:ExcelWorksheet) = seq { 
        let maxRow = getMaxRowNumber worksheet  
        for i in 1..maxRow do        
            let content = worksheet.Cells.[i,colIndex].Value.ToString()
            yield content
    }

    /// <summary>Gets all the values from given column in a given worksheet in a sequence. </summary>   
    /// <param name="rowIndex">The column index (starting from 1!).</param> 
    /// <param name="worksheet">The input worksheet.</param>
    /// <returns>Sequence of cell values</returns>
    let getRow rowIndex (worksheet:ExcelWorksheet) = seq { 
        let maxCol = getMaxColNumber worksheet  
        for i in 1..maxCol do        
            let content = worksheet.Cells.[rowIndex,i].Value.ToString()
            yield content
    }

    /// <summary>Creates a new Excel document. Save method needs to be called to write the file do disk </summary>       
    /// <param name="filename">The document filename.</param>
    /// <returns>Excel document</returns>
    let createDocument filename = 
        if (File.Exists(filename)) then
            File.Delete(filename)

        let newFile = new FileInfo(filename)        
        let package = new ExcelPackage(newFile)        
        package

    /// <summary>Adds a worksheet with a specified name to the Excel document</summary>      
    /// <param name="sheetName">Name of the inserted worksheet.</param> 
    /// <param name="document">The document.</param>
    /// <returns>Excel worksheet</returns>
    let addWorksheet sheetName (document:ExcelPackage) =
        document.Workbook.Worksheets.Add(sheetName) |> ignore
        let worksheet = document.Workbook.Worksheets.[document.Workbook.Worksheets.Count]
        worksheet.Name <- sheetName; 
        worksheet

    /// <summary>Adds a sequence as a row in to a worksheet </summary>       
    /// <param name="rownIndex">The row index.</param>
    /// <param name="worksheet">The input worksheet.</param>
    /// <param name="seq">The input sequence.</param>
    /// <returns>unit</returns>
    let addRow rowIndex (worksheet:ExcelWorksheet) seq =
        seq
        |> Seq.iteri (fun i x ->
            worksheet.Cells.[rowIndex,i+1].Value <- x
            )

    /// <summary>Adds a sequence as a column in to a worksheet </summary>       
    /// <param name="colIndex">The column index.</param>
    /// <param name="worksheet">The input worksheet.</param>
    /// <param name="seq">The input sequence.</param>
    /// <returns>unit</returns>
    let addColumn colIndex (worksheet:ExcelWorksheet) seq =
        seq
        |> Seq.iteri (fun i x ->
            worksheet.Cells.[i+1,colIndex].Value <- x
            )
            