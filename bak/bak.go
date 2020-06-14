package bak

/*func readOldExcel(file os.FileInfo) {
	f, err := excelize.OpenFile(getCurrentDirectory() + "/" + path_src + "/" + file.Name())
	if err != nil {
		return
	}
	var dataV scanDataV

	m := f.GetSheetMap()
	for n := range m {
		sheetName := f.GetSheetName(n)
		prefix := ""
		sheetNameTags := strings.Split(sheetName, "-")
		if len(sheetNameTags) > 1 {
			is, _ := isInt(sheetNameTags[0])
			if !is {
				tolog(sheetName + "不符合规则，跳过")
				continue
			}
			tolog("读取 " + sheetName)
			prefix = sheetNameTags[0]
		}else{
			tolog(sheetName + "不符合规则，跳过")
			continue
		}


		rows, _ := f.GetRows(sheetName)
		scan := false
		startCol := 0
		startRow := 0
		for r, row := range rows {
			for c, cell := range row {
				if len(cell) > 0 && strings.HasPrefix(strings.ToLower(cell), "start-") {
					// 起始扫描点
					starts := strings.Split(cell, "-")
					if len(starts) > 1 {
						scan = true
					}
					startCol = c
					startRow = r
					dataV = scanDataV{}
					dataV.startNum = starts[1]
					continue
				}

				if len(cell) > 0 && len(dataV.startNum) > 0 && strings.HasPrefix(strings.ToLower(cell), "end-"+dataV.startNum) {
					// 终止扫描点
					scan = false
					// 将已经记录的数据进行截取。得到实际的数据区间
					scanMap[prefix + "-start-"+dataV.startNum] = scanDataV{
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
					if "-" == cell {
						cell = string(0)
					}
					dataV.data[r-startRow][c-startCol-1] = NumberFormat(cell)
				}
			}
		}
	}
}
*/
