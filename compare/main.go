package main

import (
	"bufio"
	"fmt"
	"github.com/unidoc/unioffice/spreadsheet"
	"office-go/common"
	"os"
	"strconv"
	"strings"
)

var checkWork bool

const path_src = "src"

var currentDir string

var cols = [26]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func findCol(col int) string {
	if col < 26 {
		return cols[col]
	}
	index := col / 26
	first := cols[index-1]
	num := col % 26
	second := cols[num]
	return first + second
}

func main() {
	b, _ := common.AuthCheck()
	if !b {
		return
	}

	currentDir = common.GetCurrentDirectory()
	common.Log("开始 " + currentDir)

	// 读取excel列表
	excelfiles, _ := common.FindFiles(currentDir+"/"+path_src, "XLSX")
	if excelfiles == nil || len(excelfiles) == 0 {
		common.Log("src目录没有找到xlsx文件！")
		return
	}
	if len(excelfiles) != 2 {
		common.Log("xlsx比较文件不应该是有2个么？")
		return
	}

	wb1, _ := spreadsheet.Open(currentDir + "/" + path_src + "/" + excelfiles[0].Name())
	wb2, _ := spreadsheet.Open(currentDir + "/" + path_src + "/" + excelfiles[1].Name())

	sheets := wb1.Sheets()
	for _, sheet1 := range sheets {
		common.Log("开始读取sheet " + sheet1.Name())
		sheet2, er := wb2.GetSheet(sheet1.Name())
		if er != nil {
			common.Log("读取sheet失败！" + sheet1.Name() + " 未找到！")
			continue
		}
		compare1(sheet1, sheet2)
	}

	//写入日志文件
	if checkWork {
		common.Log(" excel中有不匹配的数据，请查看日志！")
	}
	common.Log("结束 ")
	_ = common.LogClose
	fmt.Printf("取数结束，按回车关闭此窗口... ")
	inputReader := bufio.NewReader(os.Stdin)
	_, _ = inputReader.ReadString('\n')
}

func compare1(sheet1 spreadsheet.Sheet, sheet2 spreadsheet.Sheet) {
	for i := 0; i < 1000; i++ {
		for j := 0; j < 100; j++ {
			col := findCol(j)
			row := strconv.Itoa(i + 1)
			cell1 := sheet1.Cell(col + row).GetString()
			cell2 := sheet2.Cell(col + row).GetString()
			if len(cell1) == 0 && len(cell2) == 0 {
				continue
			}
			if strings.Compare(cell1, cell2) != 0 {
				checkWork = true
				common.LogPrintf("匹配失败 %s , src %s, target %s ,  cell %s\n", sheet1.Name(), cell1, cell2, col+row)
			}
		}
	}
}
