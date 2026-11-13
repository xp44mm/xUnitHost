/// 读文本文件：认编码 → 解码 → 切行。返回 Task<编码名 * 行数组>。
///
/// 判定顺序（宁可标明歧义，也不替人猜）：
///   ① BOM：UTF-8 / UTF-16LE / UTF-16BE / UTF-32LE / UTF-32BE，编码名带 -bom；
///   ② 无 BOM 且所有字节 < 0x80：ascii —— 它同时是合法的 UTF-8 和 GBK，三者解出来一样；
///   ③ 无 BOM：UTF-8 严格试解码 + GBK(936) 严格试解码：
///        · 只有 UTF-8 过       → utf-8
///        · 只有 GBK 过         → gbk
///        · 都过且文本不同      → utf-8/gbk（两种都合法，正文按 UTF-8，故 UTF-8 在前）
///        · 都不过              → utf-8/gbk 或 gbk/utf-8（坏字符少的在前，更像的那个在正文）
///
///   编码名里带「/」一律表示「还有备选，调用方得继续判断」：
///     前 = 正文实际用的那套，后 = 备选；顺序即可能性由大到小。
///     两种来源不同（都合法 / 都不合法），但都是同一件事：没法定死，只能给候选。
///   ④ 行：按 \r\n / \n / \r 切（覆盖 \r?\n，也认裸 \r），末尾换行不多带一个空行。
///
/// 异步：netstandard2.0 没有 File.ReadAllBytesAsync，用 FileStream(FileOptions.Asynchronous)
///       + ReadAsync 读字节，写成 task 表达式，直接返回 Task，不占线程池线程干等。
///
/// open OsHost 之后 FSI 一行就能调：
///         #r @"...\OsHost.dll";; open OsHost;;
///         let enc, rows = (TextFile.decodeAllLinesAsync @"C:\a.txt").Result;;
namespace OsHost

open System
open System.IO
open System.Text
open System.Threading.Tasks

[<RequireQualifiedAccess>]
module TextFile =

    // .NET Core / .NET 5+ 默认只内置 UTF-8 / UTF-16 / UTF-32 / ASCII，GBK 得挂代码页 provider
    do
        try
            Encoding.RegisterProvider CodePagesEncodingProvider.Instance
        with _ ->
            ()

    let private strictUtf8 = UTF8Encoding(false, true) :> Encoding
    let private looseUtf8 = UTF8Encoding(false, false) :> Encoding

    let private strictGbk () =
        Encoding.GetEncoding(936, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)

    /// 宽松解码：坏字节换成 U+FFFD。
    /// 必须两边都换 U+FFFD —— 默认的 DecoderFallback.ReplacementFallback 换的是 '?'(U+003F)，
    /// 那样 replacementCount 只数 U+FFFD 就永远给 GBK 记 0 个坏字符，比较直接失效。
    let private looseGbk () =
        Encoding.GetEncoding(936, EncoderFallback.ReplacementFallback, DecoderReplacementFallback "\uFFFD")

    let private startsWith (prefix: byte[]) (bytes: byte[]) =
        if bytes.Length < prefix.Length then
            false
        else
            let mutable same = true
            let mutable i = 0

            while same && i < prefix.Length do
                if bytes.[i] <> prefix.[i] then
                    same <- false

                i <- i + 1

            same

    let private tryDecode (enc: Encoding) (bytes: byte[]) : string option =
        try
            Some(enc.GetString bytes)
        with :? DecoderFallbackException ->
            None

    let private replacementCount (text: string) =
        text |> Seq.sumBy(fun c -> if c = '\uFFFD' then 1 else 0)

    /// 编码名 + 用来解码的 Encoding + 要跳过的 BOM 字节数
    let private detect (bytes: byte[]) : string * Encoding * int =
        if startsWith [| 0x00uy; 0x00uy; 0xFEuy; 0xFFuy |] bytes then
            "utf-32be-bom", (UTF32Encoding(true, false) :> Encoding), 4
        elif startsWith [| 0xFFuy; 0xFEuy; 0x00uy; 0x00uy |] bytes then
            "utf-32le-bom", (UTF32Encoding(false, false) :> Encoding), 4
        elif startsWith [| 0xEFuy; 0xBBuy; 0xBFuy |] bytes then
            "utf-8-bom", (UTF8Encoding(false) :> Encoding), 3
        elif startsWith [| 0xFEuy; 0xFFuy |] bytes then
            "utf-16be-bom", (UnicodeEncoding(true, false) :> Encoding), 2
        elif startsWith [| 0xFFuy; 0xFEuy |] bytes then
            "utf-16le-bom", (UnicodeEncoding(false, false) :> Encoding), 2
        elif bytes |> Array.forall(fun b -> b < 0x80uy) then
            "ascii", Encoding.ASCII, 0
        else
            match tryDecode strictUtf8 bytes, tryDecode (strictGbk ()) bytes with
            | Some u, Some g when u <> g -> "utf-8/gbk", strictUtf8, 0
            | Some _, _ -> "utf-8", strictUtf8, 0
            | None, Some _ -> "gbk", strictGbk (), 0
            | None, None ->
                // 两种都没过：替换字符少的那种更像，排前面；平手沿用 UTF-8 优先
                let u = replacementCount (looseUtf8.GetString bytes)
                let g = replacementCount (looseGbk().GetString bytes)

                if u <= g then
                    "utf-8/gbk", looseUtf8, 0
                else
                    "gbk/utf-8", looseGbk (), 0

    /// 一次解码：编码名 + 全文（BOM 已剥掉，不会粘在第一行行首）
    let decodeAllText (bytes: byte[]) : string * string =
        let name, enc, skip = detect bytes
        name, enc.GetString(bytes, skip, bytes.Length - skip)

    let private splitLines (text: string) : string[] =
        if String.IsNullOrEmpty text then
            [||]
        else
            let parts = text.Split([| "\r\n"; "\n"; "\r" |], StringSplitOptions.None)

            if (text.EndsWith "\n" || text.EndsWith "\r") && parts.[parts.Length - 1] = "" then
                parts.[0 .. parts.Length - 2]
            else
                parts

    /// 同步底座：字节 → (编码名, 行数组)
    let decodeAllLines (bytes: byte[]) : string * string[] =
        let name, text = decodeAllText bytes
        name, splitLines text

    let private readAllBytesAsync (path: string) : Task<byte[]> =
        task {
            use stream =
                new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    1 <<< 16,
                    FileOptions.Asynchronous
                )

            let length = int stream.Length
            let bytes = Array.zeroCreate<byte> length
            let mutable offset = 0

            while offset < length do
                let! read = stream.ReadAsync(bytes, offset, length - offset)

                if read <= 0 then
                    offset <- length
                else
                    offset <- offset + read

            return bytes
        }

    /// 编码名 + 全文（异步）。文件不存在 / 读不了，异常进 Task，不吞。
    let decodeAllTextAsync (path: string) : Task<string * string> =
        task {
            let! bytes = readAllBytesAsync path
            return decodeAllText bytes
        }

    /// 编码名 + 行数组（异步）。这就是对外那一个入口。
    let decodeAllLinesAsync (path: string) : Task<string * string[]> =
        task {
            let! bytes = readAllBytesAsync path
            return decodeAllLines bytes
        }
