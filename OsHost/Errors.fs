/// 异常归因：COM 异常外层多是无信息量的 "Exception from HRESULT: 0x..."，有用的在里层
[<RequireQualifiedAccess>]
module OsHost.Errors

open System

let explain (ex: exn) : string =
    let mutable e = ex

    while not(isNull e.InnerException) do
        e <- e.InnerException

    if String.IsNullOrWhiteSpace e.Message then
        sprintf "%s (0x%X)" (e.GetType().Name) e.HResult
    else
        e.Message
