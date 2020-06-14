package main

import "office-go/common"

const path_src = "src"
const path_result = "result"

var currentDir string

type scanDataV struct {
	row      int
	col      int
	data     [500][500]string
	startNum string
}

var scanMap map[string]scanDataV

func main() {
	b, _ := common.AuthCheck()
	if !b {
		return
	}

	currentDir = common.GetCurrentDirectory()
	common.Log("开始 " + currentDir)
	_, _ = common.PathExists(currentDir+"/"+path_result, true)

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

	scanMap = make(map[string]scanDataV)

}
