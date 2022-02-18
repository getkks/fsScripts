#load "./Folder.fsx"

open System.IO
open System.Linq

let excelFileExtensions = [| ".xls" ; ".xlsb" ; ".xlsx" ; ".xlsm" ; ".csv" |].ToHashSet()

let excelFiles folderName =
    folderName
    |> Folder.Files
    |> Seq.filter (fun file ->
        file.Name
        |> Path.GetExtension
        |> excelFileExtensions.Contains
    )
