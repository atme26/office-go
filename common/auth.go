package common

import (
	"bufio"
	"encoding/base64"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"
)

func AuthCheck() (bool, error) {
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
		org, err := DesDecrypt(decodeBytes, key)
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

func isInt(value string) (bool, int64) {
	n, err := strconv.ParseInt(value, 10, 0)
	if err != nil {
		return false, 0
	}
	return true, n
}
