open System.IO

let defaultEnumerationOptions = EnumerationOptions()
defaultEnumerationOptions.IgnoreInaccessible <- true
defaultEnumerationOptions.MatchCasing <- MatchCasing.CaseInsensitive
defaultEnumerationOptions.RecurseSubdirectories <- true

let Files folderName =
    (folderName |> DirectoryInfo)
        .EnumerateFiles("*", defaultEnumerationOptions)
