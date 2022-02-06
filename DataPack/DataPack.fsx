#nowarn "760"

#r "nuget: EPPlus, 5.8.5"
#r "nuget: Spectre.Console, 0.43.0"

open System
open System.IO
open System.IO.Compression
open System.Text
open OfficeOpenXml
open Spectre.Console

let defaultOutputFormat = ExcelOutputTextFormat()
defaultOutputFormat.Encoding <- Encoding.UTF8
defaultOutputFormat.FirstRowIsHeader <- true
defaultOutputFormat.TextQualifier <- '"'

let prependCheckMark str = Emoji.Known.CheckMark + " " + str

let rangeToZipFile (range: ExcelRangeBase) fileName (zipFile: ZipArchive) =
    use writer =
        zipFile
            .CreateEntry(fileName, System.IO.Compression.CompressionLevel.Fastest)
            .Open()

    range.SaveToText(writer, defaultOutputFormat)

let packFile (ctx: LiveDisplayContext) (root: Tree) fileName =
    use outFile =
        new FileStream(Path.ChangeExtension(fileName, "zip"), FileMode.Create)

    use zipFile =
        new ZipArchive(outFile, ZipArchiveMode.Create)

    if String.Equals(fileName |> Path.GetExtension, ".csv", StringComparison.OrdinalIgnoreCase) then
        zipFile.CreateEntryFromFile(
            fileName,
            Path.ChangeExtension(fileName, ".csv")
            |> Path.GetFileName,
            System.IO.Compression.CompressionLevel.Fastest
        )
        |> ignore
    else
        use excelFile = fileName |> FileInfo |> ExcelPackage

        for sheet in excelFile.Workbook.Worksheets do
            let sheetNode =
                sheet.Name + " <-> " + sheet.Dimension.Address
                |> prependCheckMark
                |> root.AddNode

            if sheet.Tables.Count = 0 then
                zipFile
                |> rangeToZipFile sheet.Cells.[sheet.Dimension.Address] (sheet.Name + ".csv")
            else
                for table in sheet.Tables do
                    table.Name + " <-> " + table.Address.Address
                    |> prependCheckMark
                    |> sheetNode.AddNode
                    |> ignore

                    zipFile
                    |> rangeToZipFile table.Range (table.Name + ".csv")

            ctx.Refresh()
