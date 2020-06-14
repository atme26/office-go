// Copyright 2017 FoxyUtils ehf. All rights reserved.
package main

import (
	"bufio"
	"fmt"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/spreadsheet"
	"office-go/common"
	"os"
	"strings"
)

const path_src = "src"
const path_target = "target"
const path_result = "result"

var checkWork bool
var scanMap map[string]scanDataV

var tableMap map[string]tableDataV

var currentDir string

type scanDataV struct {
	row      int
	col      int
	data     [500][500]string
	startNum string
}

type tableDataV struct {
	tableNo  int
	startRow int
	startCol int
}

func main() {

	b, _ := common.AuthCheck()
	if !b {
		return
	}

	currentDir = common.GetCurrentDirectory()
	tolog("开始 " + currentDir)
	_, _ = common.PathExists(currentDir+"/"+path_result, true)

	//target
	docfile, _ := common.FindFile(currentDir+"/"+path_target, "DOCX")
	if docfile == nil {
		tolog("target目录没有找到docx文件！")
		return
	}
	// 读取excel列表
	excelfiles, _ := common.FindFiles(currentDir+"/"+path_src, "XLSX")
	if excelfiles == nil || len(excelfiles) == 0 {
		tolog("src目录没有找到xlsx文件！")
		return
	}

	scanMap = make(map[string]scanDataV)
	tableMap = make(map[string]tableDataV)
	checkWork = true

	for _, file := range excelfiles {
		readExcel(file)
	}

	if scanMap == nil {
		tolog("src中没有找到需要处理的数据！")
		return
	}

	doc := matchWord(docfile)

	err := writeToWord(doc)
	if err != nil {
		tolog(err)
	}
	_ = doc.SaveToFile(currentDir + "/" + path_result + "/temp_0_" + docfile.Name())

	_ = cleanWord(doc)
	_ = doc.SaveToFile(currentDir + "/" + path_result + "/temp_1_" + docfile.Name())

	//写入日志文件
	if !checkWork {
		tolog(" word中有校验未通过，请修改完再重试！")
	}
	tolog("结束 ")
	_ = common.LogClose
	fmt.Printf("取数结束，按回车关闭此窗口... ")
	inputReader := bufio.NewReader(os.Stdin)
	_, _ = inputReader.ReadString('\n')

}

func readExcel(file os.FileInfo) {
	wb, err := spreadsheet.Open(currentDir + "/" + path_src + "/" + file.Name())
	if err != nil {
		tolog(err)
		return
	}

	var dataV scanDataV
	sheets := wb.Sheets()
	for _, sheet := range sheets {

		scan := false
		startCol := 0
		startRow := 0
		rows := sheet.Rows()

		for r, row := range rows {
			cells := row.Cells()
			for c, cell := range cells {

				if !cell.IsEmpty() && strings.HasPrefix(strings.ToLower(cell.GetString()), "start-") {
					// 起始扫描点
					starts := strings.Split(cell.GetString(), "-")
					if len(starts) > 1 {
						scan = true
					}
					startCol = c
					startRow = r
					dataV = scanDataV{}
					dataV.startNum = starts[len(starts)-1]
					continue
				}

				/*if "7-start-15" != prefix + "-start-" + dataV.startNum {
					continue
				}
				*/
				if !cell.IsEmpty() && len(dataV.startNum) > 0 && strings.HasPrefix(strings.ToLower(cell.GetString()), "end-"+dataV.startNum) {
					// 终止扫描点
					scan = false

					//tag := prefix + "-start-" + dataV.startNum

					// 将已经记录的数据进行截取。得到实际的数据区间
					scanMap["start-"+dataV.startNum] = scanDataV{
						startNum: dataV.startNum,
						data:     dataV.data,
						row:      r - startRow + 1,
						col:      c - startCol - 1,
					}
					continue
				}
				if c <= startCol {
					continue
				}

				if c-startCol >= 40 {
					continue
				}

				if scan {
					cVell := cell.GetFormattedValue()
					cVell = strings.TrimSpace(cVell)
					cVell = strings.Trim(cVell, ",")
					if strings.HasPrefix(cVell, "(,") {
						cVell = strings.Replace(cVell, "(,", "(", -1)
					}
					dataV.data[r-startRow][c-startCol-1] = cVell
				}
			}

		}

	}
}

func matchWord(file os.FileInfo) (doc *document.Document) {
	doc, _ = document.Open(currentDir + "/" + path_target + "/" + file.Name())
	tables := doc.Tables()
	for t, table := range tables {
		startCol := 0
		startRow := 0
		scan := false
		var dataV scanDataV
		startText := "SSSSSSSSSSS"
		endText := "EEEEEEEEEE"
		rows := table.Rows()
		for r, row := range rows {
			cells := row.Cells()
			for c, cell := range cells {
				if len(cell.Paragraphs()) == 0 || len(cell.Paragraphs()[0].Runs()) == 0 || len(cell.Paragraphs()[0].Runs()) == 0 {
					continue
				}

				text := getCellValue(cell)
				if scan {
					//tolog(" 当前值 " + text)
				}
				if strings.Contains(text, "start-") {

					if scan {
						tolog(" 预期应该查找值 " + endText + "， 实际匹配值 " + text)
						checkWork = false
					}

					starts := strings.Split(text, "-")
					if len(starts) > 1 {
						scan = true
					}
					startText = starts[len(starts)-2] + "-" + starts[len(starts)-1]
					tolog("开始匹配 " + text)
					endText = "end-" + starts[len(starts)-1]
					startRow = r
					startCol = c

					break
				}

				if strings.Contains(text, "end-") {
					if !scan {
						tolog(" 忽略值 " + text)
					}

					scan = false
					if strings.Contains(text, endText) {

						rs := r - startRow - 1
						cs := c - startCol + 1
						// 从excel map 取
						dataV = scanMap[startText]

						if rs == dataV.row && cs == dataV.col {
							tolog("匹配成功 " + text)
							tableMap[startText] = tableDataV{
								startRow: startRow + 1,
								startCol: startCol,
								tableNo:  t,
							}
						} else {
							common.LogPrintf("匹配失败 %s , excel row %d , excel col %d , word row %d , word col %d\n", text, dataV.row, dataV.col, rs, cs)
							checkWork = false
						}

					} else {
						tolog(" 预期应该查找值 " + endText + "， 实际匹配值 " + text)
						checkWork = false
					}
				}

			}

		}
	}
	return doc
}

func writeToWord(doc *document.Document) error {
	tables := doc.Tables()
	// 根据tableMap 找对应的起始单元格。然后从 scanMap 获取数据
	for k, v := range tableMap {
		tableNo := tableMap[k].tableNo
		/*if k != "7-start-15" {
			continue
		}*/
		if tableNo >= len(tables) {
			tolog(k + " 对应的表格位置格式错误，跳过")
			continue
		}

		table := tables[tableNo]
		data := scanMap[k]
		if data.row == 0 && data.col == 0 {
			tolog(k + " 对应的excel没有数据，跳过")
			continue
		}

		r := v.startRow
		c := v.startCol
		i := 0
		j := 0
		for i = r; i < r+data.row; i++ {
			len1 := len(table.Rows()[i].Cells())
			for j = c; j < c+data.col; j++ {
				if j >= len1 {
					if data.data[i-r][j-c] == "" {
						continue
					} else {
						common.LogPrintf("excel 位置 %s, row %d , col %d 对应的word单元格可能存在跨列，请检查此单元格所在的行是否需要手动调整 \n", k, i-r, j-c)
						break
					}
				}
				setCellValue(table.Rows()[i].Cells()[j], data.data[i-r][j-c])
			}
		}

	}
	return nil
}

func cleanWord(doc *document.Document) error {
	tables := doc.Tables()
	// 根据tableMap 找对应的起始单元格。然后从 scanMap 获取数据
	for k, v := range tableMap {
		tableNo := tableMap[k].tableNo

		if tableNo >= len(tables) {
			tolog(k + " 对应的表格位置格式错误，跳过")
			continue
		}

		table := tables[tableNo]
		data := scanMap[k]

		startr := v.startRow - 1
		startc := v.startCol
		startText := getCellValue(table.Rows()[startr].Cells()[startc])
		if strings.Contains(startText, "start-") {
			setCellValue(table.Rows()[startr].Cells()[startc], "")
		}

		endr := v.startRow + data.row
		endc := v.startCol + data.col - 1
		endText := getCellValue(table.Rows()[endr].Cells()[endc])
		if strings.Contains(endText, "end-") {
			setCellValue(table.Rows()[endr].Cells()[endc], "")
		}
	}
	return nil
}

func setCellValue(cell document.Cell, v string) {
	ps := cell.Paragraphs()

	for _, p := range ps {
		rs := p.Runs()
		for _, r := range rs {
			p.RemoveRun(r)
		}
	}

	if len(ps) > 0 {
		runs := ps[0].Runs()
		if len(runs) > 0 {
			runs[0].Clear()
			runs[0].Properties().SetFontFamily("Arial")
			runs[0].Properties().SetSize(11)
			runs[0].AddText(v)
			runs[0].AddText(v)
		} else {
			run := ps[0].AddRun()
			run.Properties().SetFontFamily("Arial")
			run.Properties().SetSize(11)
			run.AddText(v)
		}
	} else {
		run := cell.AddParagraph().AddRun()
		run.Properties().SetFontFamily("Arial")
		run.Properties().SetSize(11)
		run.AddText(v)
	}

}
func getCellValue(cell document.Cell) string {
	v := ""
	for _, p := range cell.Paragraphs() {
		for _, run := range p.Runs() {
			v = v + run.Text()
		}
	}
	return v
}

func tolog(a ...interface{}) {
	common.Log(a)
}

func NumberFormat(str string) string {
	length := len(str)
	if length < 4 {
		return str
	}
	arr := strings.Split(str, ".") //用小数点符号分割字符串,为数组接收
	length1 := len(arr[0])
	if length1 < 4 {
		return str
	}
	count := (length1 - 1) / 3
	for i := 0; i < count; i++ {
		arr[0] = arr[0][:length1-(i+1)*3] + "," + arr[0][length1-(i+1)*3:]
	}
	return strings.Join(arr, ".") //将一系列字符串连接为一个字符串，之间用sep来分隔。
}
