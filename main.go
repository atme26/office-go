// Copyright 2017 FoxyUtils ehf. All rights reserved.
package main

import (
	"bufio"
	"encoding/base64"
	"fmt"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/spreadsheet"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
	"wordutil/des"
)

const path_src = "src"
const path_target = "target"
const path_result = "result"

var logger *log.Logger
var logFile *os.File
var checkWork bool
var scanMap map[string]scanDataV

var tableMap map[string]tableDataV

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
	b, err := authCheck()
	if !b {
		return
	}

	logFile, _ = os.OpenFile("dataReport.log", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 666)
	logger = log.New(logFile, "", log.LstdFlags)
	tolog("=============================")
	tolog("开始 " + getCurrentDirectory())

	exists, err := PathExists(getCurrentDirectory() + "/" + path_result)
	if !exists {
		tolog("生成result目录")
		_ = os.Mkdir(getCurrentDirectory()+"/"+path_result, os.ModePerm)
	}

	//target
	tFiles, err := ioutil.ReadDir(getCurrentDirectory() + "/" + path_target)
	if err != nil {
		tolog(err)
	}
	var tfile string

	for _, file := range tFiles {
		if file.IsDir() {
			tolog(file.Name() + " 是目录，跳过不处理！")
			continue
		}

		if strings.HasSuffix(strings.ToUpper(file.Name()), "DOCX") {
			tfile = file.Name()
			tolog("加载目标表... " + file.Name())
			break
		}
	}

	// 读取excel列表
	files, err := ioutil.ReadDir(getCurrentDirectory() + "/" + path_src)
	if err != nil {
		tolog(err)
	}
	//src
	scanMap = make(map[string]scanDataV)

	tableMap = make(map[string]tableDataV)
	checkWork = true
	for _, file := range files {
		if file.IsDir() {
			tolog(file.Name() + " 是目录，跳过不处理！")
			continue
		}

		if strings.HasSuffix(strings.ToUpper(file.Name()), "XLSX") {
			tolog("正在读取... " + file.Name())
			readExcel(file)
			continue
		}
		tolog(file.Name() + " 格式不匹配，跳过不处理！")
	}

	if scanMap == nil {
		return
	}
	doc, err := document.Open(getCurrentDirectory() + "/" + path_target + "/" + tfile)
	matchWord(doc)
	if !checkWork {
		tolog(" 校验未通过，请修改完再重试！")

	} else {
		err := writeToWord(doc)
		if err != nil {
			tolog(err)
		}
		_ = doc.SaveToFile(getCurrentDirectory() + "/" + path_result + "/temp_0_" + tfile)

		_ = cleanWord(doc)
		_ = doc.SaveToFile(getCurrentDirectory() + "/" + path_result + "/temp_1_" + tfile)

	}

	//写入日志文件
	tolog("结束 ")
	tolog("=============================")
	tolog("")
	_ = logFile.Close()
	fmt.Printf("取数结束，按回车关闭此窗口... ")
	inputReader := bufio.NewReader(os.Stdin)
	_, _ = inputReader.ReadString('\n')

}

func getTestDirectory() string {
	return "G:/data/2020-01-16/test1/wordutil2020-01-191"
}
func getCurrentDir() string {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
	}
	return strings.Replace(dir, "\\", "/", -1)
}

func authCheck() (bool, error) {
	for {
		fmt.Printf("Please enter the password: ")
		inputReader := bufio.NewReader(os.Stdin)
		input, err := inputReader.ReadString('\n')
		if err != nil {
			return false, err
		}
		str := strings.Replace(input, "\n", "", -1)

		if "exit" == str || "stop" == str {
			return false, nil
		}

		key := []byte("88991128")
		decodeBytes, err := base64.StdEncoding.DecodeString(str)
		if err != nil {
			continue
		}
		org, err := des.DesDecrypt(decodeBytes, key)
		if err != nil {
			continue
		}
		pwd := string(org)
		pwds := strings.Split(pwd, "-")
		if len(pwds) > 1 && "yoyyu" == pwds[0] {

			currentTime := time.Now().Local().Unix()
			b, ts := isInt(pwds[1])
			if !b {
				return false, nil
			}

			if currentTime < ts {
				return true, nil
			}
		}
		fmt.Println("Permission denied, please try again.")
	}
}

func readExcel(file os.FileInfo) {
	wb, err := spreadsheet.Open(getCurrentDirectory() + "/" + path_src + "/" + file.Name())
	if err != nil {
		logger.Println(err)
		return
	}

	var dataV scanDataV
	sheets := wb.Sheets()
	prefix := ""
	for _, sheet := range sheets {
		sheetName := sheet.Name()
		sheetNameTags := strings.Split(sheetName, "-")
		if len(sheetNameTags) > 1 {
			is, _ := isInt(sheetNameTags[0])
			if !is {
				tolog(sheetName + "不符合规则，跳过")
				continue
			}
			tolog("读取 " + sheetName)
			prefix = sheetNameTags[0]
		} else {
			tolog(sheetName + "不符合规则，跳过")
			continue
		}
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
					dataV.startNum = starts[1]
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
					scanMap[prefix+"-start-"+dataV.startNum] = scanDataV{
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

func matchWord(doc *document.Document) {

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
				if strings.Contains(text, "-start-") {

					if scan {
						tolog(" 预期应该查找值 " + endText + "， 实际匹配值 " + text)
						checkWork = false
					}

					starts := strings.Split(text, "-")
					if len(starts) > 2 {
						scan = true
					}
					startText = text
					tolog("开始匹配 " + startText)
					endText = starts[0] + "-end-" + starts[2]
					startRow = r
					startCol = c

					break
				}

				if strings.Contains(text, "-end-") {
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
							fmt.Printf("匹配失败 %s , excel row %d , excel col %d , word row %d , word col %d\n", text, dataV.row, dataV.col, rs, cs)
							logger.Printf("匹配失败 %s , excel row %d , excel col %d , word row %d , word col %d\n", text, dataV.row, dataV.col, rs, cs)
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
						fmt.Printf("excel 位置 %s, row %d , col %d 对应的word单元格可能存在跨列，请检查此单元格所在的行是否需要手动调整 \n", k, i-r, j-c)
						logger.Printf("excel 位置 %s, row %d , col %d 对应的word单元格可能存在跨列，请检查此单元格所在的行是否需要手动调整 \n", k, i-r, j-c)
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

func isInt(value string) (bool, int64) {
	n, err := strconv.ParseInt(value, 10, 0)
	if err != nil {
		return false, 0
	}
	return true, n
}

func tolog(a ...interface{}) {
	logger.Println(a)
	fmt.Println(a)
}
func getCurrentDirectory() string {
	//path := getTestDirectory()
	path := getCurrentDir()
	return path
}

func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
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
