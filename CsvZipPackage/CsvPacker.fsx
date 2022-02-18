#nowarn "760"
#load "../Scripts/Excel.fsx"
#load "../Scripts/Progress.fsx"
#r "nuget: Spectre.Console, 0.43.0"

open System.IO
open System.IO.Compression
open Spectre.Console

let pack (root : TreeNode) (zipFile : ZipArchive) fileName =
    zipFile.CreateEntryFromFile(fileName, Path.GetFileName fileName, CompressionLevel.SmallestSize)
    |> ignore
