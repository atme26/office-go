package common

import (
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
)

func GetCurrentDirectory() string {
	path := getTestDirectory()
	//path := getCurrentDir()
	return path
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

func PathExists(path string, create bool) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		if create {
			e := os.Mkdir(path, os.ModePerm)
			if e != nil {
				return false, e
			}
			return true, nil
		}
		return false, nil
	}
	return false, err
}

// 从目录中查找 指定格式的文件
func FindFiles(path string, format string) ([]os.FileInfo, error) {
	tFiles, err := ioutil.ReadDir(path)
	if err != nil {
		Log(err)
	}
	fi := make([]os.FileInfo, 0)
	for _, file := range tFiles {
		if file.IsDir() {
			Log(file.Name() + " 是目录，跳过不处理！")
			continue
		}

		if strings.HasSuffix(strings.ToUpper(file.Name()), format) {
			Log("加载目标表... " + file.Name())
			fi = append(fi, file)
		}
	}
	return fi, nil
}

func FindFile(path string, format string) (os.FileInfo, error) {
	tFiles, err := ioutil.ReadDir(path)
	if err != nil {
		Log(err)
	}
	for _, file := range tFiles {
		if file.IsDir() {
			Log(file.Name() + " 是目录，跳过不处理！")
			continue
		}

		if strings.HasSuffix(strings.ToUpper(file.Name()), format) {
			Log("加载目标表... " + file.Name())
			return file, nil
		}
	}
	return nil, nil
}
