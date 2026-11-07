module internal AcadMgdHost.AssemblyInfo

open Autodesk.AutoCAD.Runtime

// NETLOAD 之后让 AutoCAD 调 Entry.Initialize()。
// F# 的程序集级属性必须落在编译顺序最后一个文件里的 do () 之前。
[<assembly: ExtensionApplication(typeof<AcadMgdHost.Entry>)>]
// 显式声明命令类：不靠控制台去扫描整个程序集找 CommandMethod。
// 实测 AutoCAD 2020 的 accoreconsole 不声明也能注册；留着更明确。
[<assembly: CommandClass(typeof<AcadMgdHost.Commands>)>]
do ()
