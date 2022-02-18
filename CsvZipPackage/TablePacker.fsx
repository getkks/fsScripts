#nowarn "760"
#load "../Scripts/Excel.fsx"
#load "../Scripts/Progress.fsx"
#r "nuget: Humanizer.Core, 2.14.1"
#r "nuget: Spectre.Console, 0.43.0"
#r "nuget: NPOI, 2.5.5"
#time "on"

open System.IO
open System.IO.Compression
open System.Text
open Spectre.Console
open NPOI.HSSF.UserModel
open NPOI.XSSF.UserModel
open NPOI.SS.UserModel
open NPOI.SS.Util
open Humanizer
open System

let private folder =
    fsi.CommandLineArgs
    |> Array.tail
    |> Array.head

let treeNode name (range : CellRangeAddress option) sheetNode =
    if range.IsNone then
        name
    else
        $"{name} - [cyan3]{range.Value.FormatAsString()}[/]"
    |> sheetNode.AddNode

let rec writeCellR (writer : StreamWriter) (cell : ICell) (cellType : CellType) =
    match cellType with
    | CellType.Blank
    | CellType.Error -> ()
    | CellType.Formula -> writeCellR writer cell cell.CachedFormulaResultType
    | CellType.Boolean -> writer.Write(cell.BooleanCellValue)
    | CellType.String ->
        let mutable value = cell.StringCellValue

        if value.IndexOfAny([| ',' ; '"' |]) <> -1 then
            value <-
                "\""
                + value.Replace("\"", "\"\"")
                + "\""

        writer.Write(value)
    | _ -> writer.Write(cell)

let writeCell (writer : StreamWriter) (cell : ICell) = writeCellR writer cell cell.CellType

let rangeToZipFileNPOI (sheet : ISheet) (range : CellRangeAddress option) fileName (zipFile : ZipArchive) =
    use file = zipFile.CreateEntry(fileName, System.IO.Compression.CompressionLevel.Fastest).Open()
    use writer = StreamWriter(file, Encoding.UTF8)

    match range with
    | None ->
        let loopCells (row : IRow) =
            use cellEnumerator = row.GetEnumerator()

            if cellEnumerator.MoveNext() then
                cellEnumerator.Current
                |> writeCell writer

                while cellEnumerator.MoveNext() do
                    writer.Write(",")

                    cellEnumerator.Current
                    |> writeCell writer

        let rowEnumerator = sheet.GetRowEnumerator()

        if rowEnumerator.MoveNext() then
            loopCells (rowEnumerator.Current :?> IRow)

            while rowEnumerator.MoveNext() do
                writer.WriteLine()
                loopCells (rowEnumerator.Current :?> IRow)

    | Some range ->
        let loopCells (row : IRow) =
            range.FirstColumn
            |> row.GetCell
            |> writeCell writer

            for col = range.FirstColumn + 1 to range.LastColumn do
                writer.Write(",")
                col |> row.GetCell |> writeCell writer

        range.FirstRow
        |> sheet.GetRow
        |> loopCells

        for row = range.FirstRow + 1 to range.LastRow do
            writer.WriteLine()
            row |> sheet.GetRow |> loopCells

let pack (root : TreeNode) fileName =
    use outFile = new FileStream(Path.ChangeExtension(fileName, "zip"), FileMode.Create)
    use zipFile = new ZipArchive(outFile, ZipArchiveMode.Create)
    use file = new FileStream(fileName, FileMode.Open, FileAccess.Read)
    let mutable workBook = Unchecked.defaultof<_>

    match Path.GetExtension fileName with
    | ".csv" ->
        zipFile.CreateEntryFromFile(fileName, Path.GetFileName fileName, System.IO.Compression.CompressionLevel.SmallestSize)
        |> ignore
    | ".xlsb" -> ()
    | ".xls" -> workBook <- HSSFWorkbook file :> IWorkbook
    | _ -> workBook <- XSSFWorkbook file :> IWorkbook

    if workBook |> isNull |> not then
        for sheet in workBook do
            let range =
                CellRangeAddress(
                    sheet.FirstRowNum,
                    sheet.LastRowNum,
                    sheet.GetRow(sheet.FirstRowNum).FirstCellNum
                    |> int,
                    (sheet.GetRow(sheet.LastRowNum).LastCellNum
                     |> int)
                    - 1
                )
                |> Some

            let sheetNode = treeNode sheet.SheetName range root

            match sheet with
            | :? XSSFSheet as sheet when sheet.GetTables().Count > 0 ->
                let tables = sheet.GetTables()

                for table in tables do
                    let range =
                        CellRangeAddress(table.StartRowIndex, table.EndRowIndex, table.StartColIndex, table.EndColIndex)
                        |> Some

                    treeNode table.Name range sheetNode
                    |> ignore

                    rangeToZipFileNPOI sheet range (table.Name + ".csv") zipFile

            | _ -> rangeToZipFileNPOI sheet range (sheet.SheetName + ".csv") zipFile

try
    AnsiConsole.Record()

    folder
    |> Excel.excelFiles
    |> Seq.toArray
    |> Progress.progressBar
        "[bold green]Zip CSV Package[/]"
        (folder |> Tree)
        (fun tree x ->
            tree.AddNode($"{Path.GetRelativePath(folder, x.FullName)} [steelblue1]{x.Length.Bytes().Humanize()}[/]")
        )
        (fun _ treeNode x -> pack treeNode x.FullName)
        (fun tree ->
            AnsiConsole.Write tree
            let time = DateTime.Now.ToString("yyyy-MM-dd HH-mm")
            use writer = File.CreateText(Path.Join(folder, $"log {time}.html"))
            AnsiConsole.ExportHtml() |> writer.Write
        )
with
    | e -> AnsiConsole.WriteException e
