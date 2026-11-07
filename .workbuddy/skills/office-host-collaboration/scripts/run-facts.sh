#!/usr/bin/env bash
#
# 编译 + 跑 xUnitHost 的测试 + 检查残进程。三步闭环里后两步的一键版。
#
#   ./run-facts.sh -m "*关键字*"     只跑匹配的 Fact
#   ./run-facts.sh -a                全量跑
#   ./run-facts.sh -m "*x*" -n       跳过编译，直接跑
#
# 退出码沿用测试程序的退出码；跑完会对比「跑之前 / 跑之后」的进程，
# 多出来的算没关干净（会返回 1），本来就在跑的不会被误认。
#
# 带 -showLiveOutput：每个 Fact 写进 output 的报告会实时打出来，
# 那份报告就是交给用户的返回信息，不用另外再写一遍。

set -uo pipefail

method="*"
all=0
build=1

while getopts "m:anh" opt; do
    case "$opt" in
        m) method="$OPTARG" ;;
        a) all=1 ;;
        n) build=0 ;;
        h)
            sed -n '2,12p' "$0"
            exit 0
            ;;
        *)
            echo "用法：$0 [-m 匹配] [-a] [-n]" >&2
            exit 2
            ;;
    esac
done

# 从脚本所在位置往上找解决方案根（有 xUnitHost.slnx 的那层）
root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

while [ "$root" != "/" ] && [ ! -f "$root/xUnitHost.slnx" ]; do
    root="$(dirname "$root")"
done

if [ ! -f "$root/xUnitHost.slnx" ]; then
    echo "找不到 xUnitHost.slnx，脚本可能被移出工作区了" >&2
    exit 2
fi

# 跑之前先记一笔：哪些 Office 进程是「本来就在跑」的（用户手上的，别动）
snapshot() {
    ps -W 2>/dev/null | grep -Ei '(WINWORD|EXCEL|acad)[.]exe' | awk '{print $4" "$NF}' | sort
}

before_file="$(mktemp)"
after_file="$(mktemp)"
trap 'rm -f "$before_file" "$after_file"' EXIT
snapshot > "$before_file"

if [ "$build" = 1 ]; then
    echo "编译：$root"
    dotnet build "$root/xUnitHost.slnx" -v q
    if [ $? -ne 0 ]; then
        echo "编译失败，先修代码" >&2
        exit 1
    fi
fi

exe="$root/xUnitHost/bin/Debug/net10.0/xUnitHost.exe"

if [ ! -f "$exe" ]; then
    echo "找不到 $exe（先编译一次）" >&2
    exit 2
fi

if [ "$all" = 1 ]; then
    echo "全量：Word / Excel 会被拉起来（AutoCAD 只在本来开着时参与，"
    echo "      没开就跳过；真要测它，先 export XUNITHOST_LAUNCH_ACAD=1）"
    "$exe" -noColor -showLiveOutput
else
    "$exe" -method "$method" -noColor -showLiveOutput
fi

code=$?

snapshot > "$after_file"

echo ""
echo "---- 现场 ----"

pre="$(cat "$before_file")"
new="$(comm -13 "$before_file" "$after_file" 2>/dev/null)"

if [ -n "$pre" ]; then
    echo "开始前就开着（你手上的，测试没碰）："
    echo "$pre" | sed 's/^/  /'
fi

if [ -z "$new" ]; then
    echo "跑完没多出任何进程：收尾干净"
else
    echo "跑完多出来的（该关的没关干净）："
    echo "$new" | sed 's/^/  /'
    code=1
fi

exit $code
