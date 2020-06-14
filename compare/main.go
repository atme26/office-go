package main

import (
	"bufio"
	"fmt"
	"github.com/unidoc/unioffice/spreadsheet"
	"office-go/common"
	"os"
	"strings"
)

var checkWork bool

const path_src = "src"

var currentDir string

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
		d1 := getCells(sheet1)
		d2 := getCells(sheet2)
		compare(sheet1.Name(), d1, d2)
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

func compare(sheet string, strings1 [1000][100]string, strings2 [1000][100]string) {
	for i := 0; i < 1000; i++ {
		for j := 0; j < 100; j++ {
			if len(strings1[i][j]) == 0 && len(strings2[i][j]) == 0 {
				continue
			}
			if strings.Compare(strings1[i][j], strings2[i][j]) != 0 {
				checkWork = true
				common.LogPrintf("匹配失败 %s , src %s, target %s  row %d ,  col %d\n", sheet, strings1[i][j], strings2[i][j], i, j)
			}
		}
	}
}

func getCells(sheet spreadsheet.Sheet) [1000][100]string {
	rows := sheet.Rows()
	var d1 [1000][100]string
	for r, row := range rows {
		if r > 999 {
			break
		}
		cells := row.Cells()
		for c, cell := range cells {
			if c > 99 {
				break
			}
			if cell.IsEmpty() {
				d1[r][c] = ""
			} else {
				d1[r][c] = cell.GetString()
			}
		}
	}
	return d1
}
