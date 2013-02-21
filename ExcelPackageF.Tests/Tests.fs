namespace ExcelPackageF.Tests

open System
open NUnit.Framework
open ExcelPackageF

[<TestFixture>]
type ``Test reading Excel files`` () =
    let worksheet = 
        @"SimpleTest.xlsx"
        |> Excel.getWorksheetByIndex 1    

    [<Test>]
    member x.``Worksheet should load without problems`` () =
        Assert.IsNotNull(worksheet)

    [<Test>]
    member x.``Worksheet should have 3 rows`` () = 
        let maxRowIndex = Excel.getMaxRowNumber worksheet
        Assert.AreEqual(maxRowIndex,3)

    [<Test>]
    member x.``Worksheet should have 2 columns`` () = 
        let maxColIndex = Excel.getMaxColNumber worksheet
        Assert.AreEqual(maxColIndex,2)

    [<Test>]
    member x.``Third row should be equal to (x,y)`` () = 
        let row = 
            worksheet
            |> Excel.getRow 3
            |> List.ofSeq

        Assert.AreEqual(row.Length,2)
        Assert.AreEqual(row,["x";"y"])

    [<Test>]
    member x.``Second columns should be equal to (b,2,y)`` () = 
        let col = 
            worksheet
            |> Excel.getColumn 2
            |> List.ofSeq

        Assert.AreEqual(col.Length,3)
        Assert.AreEqual(col,["b";"2";"y"])
    
    

[<TestFixture>]
type ``Test writing Excel files`` () =
    
    [<Test>]
    member x.``After adding two worksheet the count should be two`` () = 
        let newDocument =
            @"NewWorksheet.xlsx"
            |> Excel.createDocument

        newDocument
        |> Excel.addWorksheet "Sheet A"
        |> ignore

        Assert.AreEqual(newDocument.Workbook.Worksheets.Count,1)

        newDocument
        |> Excel.addWorksheet "Sheet B"
        |> ignore

        Assert.AreEqual(newDocument.Workbook.Worksheets.Count,2)

    [<Test>]
    member x.``Added row should match after being read back`` () = 
        let newDocument =
            @"NewWorksheet2.xlsx"
            |> Excel.createDocument

        let newSheet = 
            newDocument
            |> Excel.addWorksheet "Test sheet"

        Assert.AreEqual(newDocument.Workbook.Worksheets.Count,1)

        ["a";"b";"c";"d"]
            |> Excel.addRow 1 newSheet            

        let row = 
            newSheet
            |> Excel.getRow 1
            |> List.ofSeq

        Assert.AreEqual(row.Length,4)
        Assert.AreEqual(row,["a";"b";"c";"d"])

    [<Test>]
    member x.``Added column should match after being read back`` () = 
        let newDocument =
            @"NewWorksheet3.xlsx"
            |> Excel.createDocument

        let newSheet = 
            newDocument
            |> Excel.addWorksheet "Test sheet 2"

        Assert.AreEqual(newDocument.Workbook.Worksheets.Count,1)

        ["1";"2";"3";"4";"5"]
            |> Excel.addColumn 1 newSheet            

        let col = 
            newSheet
            |> Excel.getColumn 1
            |> List.ofSeq

        Assert.AreEqual(col.Length,5)
        Assert.AreEqual(col,["1";"2";"3";"4";"5"])