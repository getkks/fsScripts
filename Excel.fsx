#load "DataPack/DataPack.fsx"
#load "Folder.fsx"
#r "nuget: Spectre.Console, 0.43.0"
#time "on"

open System.IO
open System.Linq
open Spectre.Console

let excelFileExtensions =
    [| ".xlsx"; ".xlsm"; ".csv" |].ToHashSet()

let excelFiles folderName =
    folderName
    |> Folder.Files
    |> Seq.filter (fun file ->
        file.Name
        |> Path.GetExtension
        |> excelFileExtensions.Contains)

__SOURCE_DIRECTORY__
|> excelFiles
|> Seq.iter (fun file ->
    let tree = file.FullName |> Tree

    (tree |> AnsiConsole.Live)
        .Start(fun ctx -> file.FullName |> DataPack.packFile ctx tree))
