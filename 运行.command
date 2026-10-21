#!/bin/bash
# 双击运行：合并 Excel 文件
# 作用：自动选用正确的 Python，交互式问你要合并哪个文件夹、用哪种方式

cd "$(dirname "$0")" || exit 1

# 优先用已装好依赖的 Python；找不到就退回系统 python3
PY="/Users/wang/.workbuddy/binaries/python/versions/3.13.12/bin/python3"
if [ ! -x "$PY" ]; then
  PY="$(command -v python3)"
fi

if [ -z "$PY" ]; then
  echo "没找到 python3，请先安装 Python 3.11 或更高版本。"
  echo "按回车键关闭…"; read -r _
  exit 1
fi

echo "使用的 Python：$($PY --version 2>&1)"
echo
echo "================  Excel 合并工具  ================"
echo
echo "请把「要合并的文件夹」拖进这个窗口，然后按回车"
echo "（直接回车 = 合并当前目录）"
read -r INPUT_DIR
INPUT_DIR="${INPUT_DIR//\"/}"     # 去掉拖拽自动加上的引号
INPUT_DIR="${INPUT_DIR%/}"        # 去掉结尾的斜杠
[ -z "$INPUT_DIR" ] && INPUT_DIR="."

echo
echo "请选择合并方式："
echo "  1) 纵向拼接  —— 几个文件表头一样，把数据行堆起来（默认，直接回车即可）"
echo "  2) 横向合并  —— 类似 VLOOKUP，按关键列把列拼过来"
echo "  3) 多表汇总  —— 把每个工作表（sheet）汇总成一张表"
read -r CHOICE

EXTRA=""
case "$CHOICE" in
  2)
    MODE="merge"
    echo
    echo "请输入关键列名（多个用空格分开，例如：学号 姓名）："
    read -r KEYS
    EXTRA="-k $KEYS"
    ;;
  3)
    MODE="sheets"
    echo
    echo "是否连子目录一起扫描？（y/N）"
    read -r REC
    [ "$REC" = "y" ] && EXTRA="--recursive"
    ;;
  *)
    MODE="concat"
    echo
    echo "是否连子目录一起扫描？（y/N）"
    read -r REC
    [ "$REC" = "y" ] && EXTRA="--recursive"
    ;;
esac

echo
echo "------------------ 开始执行 ------------------"
"$PY" merge_excel.py -i "$INPUT_DIR" -m "$MODE" $EXTRA -o ./output
echo "----------------------------------------------"
echo
echo "结果已保存到 output 文件夹。按回车键关闭窗口…"
read -r _
