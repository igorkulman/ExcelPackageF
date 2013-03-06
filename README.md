# ExcelPackageF

ExcelPackageF is a simple F# wrapper over the [EPPlus library](http://epplus.codeplex.com/).

## Motivation

On Windows you can use COM to work with Excel but it is of no use if you do not have Excel installed or are using MacOS or Linux. The ExcelPackage library works without Excel and can be used with Mono.

## Usage

The ExcelPackageF namespace contains the following methods

Excel.getWorksheets  
Excel.getWorksheetByName   
Excel.getWorksheetByIndex   
Excel.getMaxRowNumber   
Excel.getMaxColNumber  
Excel.getContent  
Excel.getColumn  
Excel.getRow  

For example, if you want to read the whole data from sheet number 1 from a file called test.xlsx  

```
#!fsharp

let data = 
        "test.xlsx"
        |> Excel.getWorksheetByIndex 1
        |> Excel.getContent 

data 
    |> Seq.iter (fun x-> printfn "%s" x)
```