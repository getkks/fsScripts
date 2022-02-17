#r "nuget: Spectre.Console, 0.43.0"

open Spectre.Console

let private progress = AnsiConsole.Progress()
progress.AutoClear <- false
progress.AutoRefresh <- true
progress.HideCompleted <- false

progress.Columns(
    [| TaskDescriptionColumn() :> ProgressColumn
       ProgressBarColumn()
       PercentageColumn()
       ElapsedTimeColumn()
       SpinnerColumn(Spinner.Known.Ascii) |]
)

let progressBar taskName orignialState stateMangement callback closing data =
    progress.Start(fun ctx ->
        let p = ctx.AddTask(taskName, true, data |> Array.length |> float)

        data
        |> Seq.map (fun item ->
            let newState = stateMangement orignialState item

            async {
                callback orignialState newState item
                p.Increment 1
            }
        )
        |> Async.Parallel
        |> Async.Ignore
        |> Async.RunSynchronously
    )

    closing orignialState
