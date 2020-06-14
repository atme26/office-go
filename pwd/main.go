// Copyright 2017 FoxyUtils ehf. All rights reserved.
package main

import (
	"bufio"
	"encoding/base64"
	"fmt"
	"net"
	"os"
	"strconv"
	"strings"
	"time"
	"wordutil/des"
)

// replace github.com/unidoc/unioffice => C:\Users\xiajl\go\src\github.com\unidoc\unioffice
func main() {
	inputReader := bufio.NewReader(os.Stdin)
	var i int
	var timestamp int64
	fmt.Println("Date format like this: yyyy-MM-dd hh:mm:ss")
	for i = 0; i < 3; i++ {
		timestamp = getInputTimestamp(inputReader)
		if timestamp > 0 {
			break
		}
	}
	if timestamp == 0 {
		return
	}
	whiteMac := [...]string{"54:ee:75:d2:b9:d6", "‎88-B1-11-E7-B2-DD"}
	interfaces, err := net.Interfaces()
	if err != nil {
		panic("Poor soul,here is what you got: " + err.Error())
	}

	match := false
	for _, inter := range interfaces {
		if match {
			break
		}
		mac := inter.HardwareAddr //获取本机MAC地址
		for i := 0; i < len(whiteMac); i++ {
			match = strings.Compare(strings.ToUpper(mac.String()), strings.ToUpper(whiteMac[i])) == 0
			if match {
				break
			}
		}
	}

	if !match {
		fmt.Println("You not in white list")
	}

	timeStr := strconv.FormatInt(timestamp, 10)
	data := []byte("yoyyu-" + timeStr)
	key := []byte("88991128")
	pwd, _ := des.DesEncrypt(data, key)

	encodeString := base64.StdEncoding.EncodeToString(pwd)

	fmt.Println("密码已经生成：" + encodeString)
	fmt.Printf("复制密码后，直接按回车关闭此窗口... ")
	_, _ = inputReader.ReadString('\n')
}

func getInputTimestamp(reader *bufio.Reader) int64 {
	fmt.Printf("Please enter the expire date: ")
	input, err := reader.ReadString('\n')
	str := input
	str = strings.Replace(input, "\r\n", "", -1)
	str = strings.Replace(str, "\r", "", -1)
	str = strings.Replace(str, "\n", "", -1)
	loc, _ := time.LoadLocation("Asia/Shanghai")
	tt, err := time.ParseInLocation("2006-01-02 15:04:05", str, loc)
	if err != nil {
		fmt.Println("Enter date format error, example like this:", "2006-01-02 15:04:05")
		return 0
	}
	return tt.Unix()
}
