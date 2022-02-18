#nowarn "760"
#load "../Scripts/Excel.fsx"
#load "../Scripts/Progress.fsx"
#load "CsvPacker.fsx"
#r "nuget: Humanizer.Core, 2.14.1"
#r "nuget: Spectre.Console, 0.43.0"
#r "nuget: Sylvan.Data.Excel, 0.1.*"
#r "nuget: Sylvan.Data.Csv, 1.1.*"
#r "nuget: Ben.Demystifier, 0.4.1"
#time "on"

open System
open System.Diagnostics
open System.IO
open System.IO.Compression
open System.Text
open Humanizer
open Spectre.Console
open Sylvan.Data.Csv
open Sylvan.Data.Excel

let private folder =
    fsi.CommandLineArgs
    |> Array.tail
    |> Array.head

let prepandEmoji emoji str = if AnsiConsole.Profile.Capabilities.Legacy then str else $"{emoji} {str}"

let pack (root : TreeNode) fileName =
    use outFile = new FileStream(Path.ChangeExtension(fileName, "zip"), FileMode.Create)
    use zipFile = new ZipArchive(outFile, ZipArchiveMode.Create)

    match Path.GetExtension fileName with
    | ".csv" -> CsvPacker.pack root zipFile fileName
    | _ ->
        use reader = ExcelDataReader.Create fileName

        while (let sheetName = reader.WorksheetName
               let csvFileName = $"{sheetName}.csv"

               $"{csvFileName}"
               |> prepandEmoji $"[darkgreen]{Emoji.Known.CheckMark}[/]"
               |> root.AddNode
               |> ignore

               use fileEntry = zipFile.CreateEntry(csvFileName, CompressionLevel.SmallestSize).Open()
               use fileWriter = StreamWriter(fileEntry, Encoding.UTF8)
               use writer = CsvDataWriter.Create(fileWriter)
               writer.Write reader |> ignore
               reader.NextResult()) do
            ()

try
    AnsiConsole.Record()

    folder
    |> Excel.excelFiles
    |> Seq.toArray
    |> Progress.progressBar
        "[bold green]Zip CSV Package[/]"
        (folder
         |> prepandEmoji Emoji.Known.FileFolder
         |> Tree)
        (fun tree x ->
            $"{Path.GetRelativePath(folder, x.FullName)} [steelblue1]{x.Length.Bytes().Humanize()}[/]"
            |> prepandEmoji Emoji.Known.Memo
            |> tree.AddNode
        )
        (fun _ treeNode x -> pack treeNode x.FullName)
        (fun tree ->
            AnsiConsole.Write tree
            let time = DateTime.Now.ToString("yyyy-MM-dd HH-mm")
            use writer = File.CreateText(Path.Join(folder, $"log {time}.html"))
            AnsiConsole.ExportHtml() |> writer.Write
        )
with
    | e ->
        e.Demystify()
        |> AnsiConsole.WriteException
