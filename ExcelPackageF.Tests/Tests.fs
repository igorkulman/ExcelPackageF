﻿namespace ExcelPackageF.Tests

open System
open NUnit.Framework
open ExcelPackageF

[<TestFixture>]
type Test() =
    let worksheet = 
        @"..\..\SimpleTest.xlsx"
        |> Excel.getWorksheetByIndex 1

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
        
