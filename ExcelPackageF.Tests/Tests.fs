namespace ExcelPackageF.Tests

open System
open NUnit.Framework
open ExcelPackageF

[<TestFixture>]
type Test() =
    let worksheet = 
        @"SimpleTest.xlsx"
        |> Excel.getWorksheetByIndex 1

    let newDocument =
        @"NewWorksheet.xlsx"
        |> Excel.createDocument

    [<Test>]
    member x.LoadWorksheet () =
        Assert.IsNotNull(worksheet)

    [<Test>]
    member x.GetRowCount () = 
        let maxRowIndex = Excel.getMaxRowNumber worksheet
        Assert.AreEqual(maxRowIndex,3)

    [<Test>]
    member x.GetColCount () = 
        let maxColIndex = Excel.getMaxColNumber worksheet
        Assert.AreEqual(maxColIndex,2)

    [<Test>]
    member x.GetRow () = 
        let row = 
            worksheet
            |> Excel.getRow 3
            |> List.ofSeq

        Assert.AreEqual(row.Length,2)
        Assert.AreEqual(row,["x";"y"])

    [<Test>]
    member x.GetCol () = 
        let col = 
            worksheet
            |> Excel.getColumn 2
            |> List.ofSeq

        Assert.AreEqual(col.Length,3)
        Assert.AreEqual(col,["b";"2";"y"])
    
    member x.DeleteWorksheets () =
        newDocument.Workbook.Worksheets
        |> Seq.iter (fun x -> newDocument.Workbook.Worksheets.Delete(newDocument.Workbook.Worksheets.Count))

        Assert.AreEqual(newDocument.Workbook.Worksheets.Count,0)

    [<Test>]
    member x.AddWorksheet () = 
        x.DeleteWorksheets |> ignore

        newDocument
        |> Excel.addWorksheet "Sheet A"
        |> ignore

        Assert.AreEqual(newDocument.Workbook.Worksheets.Count,1)

        newDocument
        |> Excel.addWorksheet "Sheet B"
        |> ignore

        Assert.AreEqual(newDocument.Workbook.Worksheets.Count,2)

    [<Test>]
    member x.AddRow () = 
        x.DeleteWorksheets |> ignore

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
    member x.AddCol () = 
        x.DeleteWorksheets |> ignore

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

    [<Test>]
    member x.Save () =
        newDocument.Save()
