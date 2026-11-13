/// 时间：地址里一律用 Unix 毫秒（跨进程、跨机器都没有歧义），报告里才换成人能读的。
[<RequireQualifiedAccess>]
module OsHost.Clock

open System

let private epochTicks = DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks

let unixMs (t: DateTime) : int64 = (t.ToUniversalTime().Ticks - epochTicks) / 10000L

/// 0 就是「还没保存过」，不摆一个 1970 出来
let show (ms: int64) : string =
    if ms <= 0L then
        "还没保存过"
    else
        DateTimeOffset.FromUnixTimeMilliseconds(ms).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss.fff")

