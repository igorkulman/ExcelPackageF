namespace ExcelPackageF.Tests

open System
open NUnit.Framework
open ExcelPackageF

[<TestFixture>]
type Test() =
    let worksheet = 
        @"SimpleTest.xlsx"
        |> Excel.getWorksheetByIndex 1

    let newWorksheet =
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

    [<Test>]
    member x.AddWorksheet () = 
        newWorksheet
        |> Excel.addWorksheet "Sheet A"
        |> ignore

        Assert.AreEqual(newWorksheet.Workbook.Worksheets.Count,1)

        newWorksheet
        |> Excel.addWorksheet "Sheet B"
        |> ignore

        Assert.AreEqual(newWorksheet.Workbook.Worksheets.Count,2)

    [<Test>]
    member x.Save () =
        newWorksheet.Save()
