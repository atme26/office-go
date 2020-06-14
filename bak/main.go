// Copyright 2017 FoxyUtils ehf. All rights reserved.
package bak

import (
	"bufio"
	"encoding/base64"
	"fmt"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/spreadsheet"
	"io/ioutil"
	"log"
	"office-go/common"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

const path_src = "src"
const path_target = "target"
const path_result = "result"

var logger *log.Logger
var logFile *os.File
var scanMap map[string]scanDataV

type scanDataV struct {
	row      int
	col      int
	data     [100][100]string
	startNum string
}

func main() {
	b, err := authCheck()
	if !b {
		return
	}

	// doc, err := document.Open("G:/data/2019-07-07/tables.docx")
	logFile, _ = os.OpenFile("dataReport.log", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 666)
	logger = log.New(logFile, "", log.LstdFlags)
	tolog("=============================")
	tolog("开始")

	exists, err := PathExists(getCurrentDirectory() + "/" + path_result)
	if !exists {
		tolog("生成result目录")
		_ = os.Mkdir(getCurrentDirectory()+"/"+path_result, os.ModePerm)
	}

	// 读取excel列表
	files, err := ioutil.ReadDir(getCurrentDirectory() + "/" + path_src)
	if err != nil {
		tolog(err)
	}
	//src
	scanMap = make(map[string]scanDataV)

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
		_ = logFile.Close()
		return
	}
	tfs, err := ioutil.ReadDir(getCurrentDirectory() + "/" + path_target)

	for _, tf := range tfs {
		if tf.IsDir() {
			tolog(tf.Name() + " 是目录，跳过不处理！")
			continue
		}

		if strings.HasSuffix(strings.ToUpper(tf.Name()), "XLSX") {
			tolog("正在写入... " + tf.Name())
			writeToTarget(tf)
			continue
		}
		tolog(tf.Name() + " 格式不匹配，跳过不处理！")
	}
	//写入日志文件
	tolog("结束 ")
	tolog("=============================")
	tolog("")
	_ = logFile.Close()

}

func getTestDirectory() string {
	return "G:/data/2020-01-16/test2"
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
		org, err := common.DesDecrypt(decodeBytes, key)
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

			fmt.Println(currentTime)
			fmt.Println(ts)
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
	for _, sheet := range sheets {
		sheetName := sheet.Name()
		tolog("读取 " + sheetName)

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
				if !cell.IsEmpty() && len(dataV.startNum) > 0 && strings.HasPrefix(strings.ToLower(cell.GetString()), "end-"+dataV.startNum) {
					// 终止扫描点
					scan = false
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
					if "-" == cell.GetString() {
						dataV.data[r-startRow][c-startCol-1] = "0"
					} else {
						cVell := cell.GetFormattedValue()
						if strings.HasPrefix(cVell, "(,") {
							cVell = strings.Replace(cVell, "(,", "(", -1)
						}
						dataV.data[r-startRow][c-startCol-1] = cVell
					}
				}
			}

		}

	}
}

func writeToTarget(file os.FileInfo) {
	if scanMap == nil {
		return
	}
	wb, err := spreadsheet.Open(getCurrentDirectory() + "/" + path_target + "/" + file.Name())
	if err != nil {
		logger.Println(err)
		return
	}

	sheets := wb.Sheets()
	for _, sheet := range sheets {
		sheetName := sheet.Name()
		tolog("读取 " + sheetName)
		match := false
		startCol := 0
		startRow := 0
		var dataV scanDataV

		rows := sheet.Rows()

		for r, row := range rows {
			cells := row.Cells()
			for c, cell := range cells {
				if !cell.IsEmpty() && strings.HasPrefix(strings.ToLower(cell.GetString()), "start-") {
					// 开始匹配
					match = true
					startCol = c
					startRow = r
					dataV = scanMap[cell.GetString()]
					continue
				}

				if c <= startCol {
					continue
				}

				if match {
					// 行范围内
					if r-startRow >= dataV.row {
						continue
					}
					//列范围内
					if c-startCol > dataV.col {
						continue
					}
					value := dataV.data[r-startRow][c-startCol-1]

					cValue := cell.GetFormattedValue()
					if value != cValue {
						cell.SetString(cValue)
						tolog(sheetName + "  start-" + dataV.startNum + " " + "  原值：" + cValue + " -->  新值： " + value)
					}
				}
			}

		}

	}
	_ = wb.SaveToFile(getCurrentDirectory() + "/" + path_result + "/temp_" + file.Name())
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
	return getTestDirectory()
	//return getCurrentDir()
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
