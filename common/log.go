package common

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"time"
)

var logger *log.Logger
var logFile *os.File

func init() {
	GetLogger()
}

func GetLogger() {
	timeUnix := time.Now().Unix()
	fileName := "dataReport-" + strconv.FormatInt(timeUnix, 10) + ".log"
	logFile, _ = os.OpenFile(fileName, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 666)
	logger = log.New(logFile, "", log.LstdFlags)
}

func Log(a ...interface{}) {
	logger.Println(a...)
	fmt.Println(a...)
}

func LogPrintf(format string, a ...interface{}) {
	logger.Printf(format, a...)
	fmt.Printf(format, a...)
}

func Logfmt(a ...interface{}) {
	fmt.Println(a...)
}

func LogfmtPrintf(format string, a ...interface{}) {
	fmt.Printf(format, a...)
}

func LogClose() {
	_ = logFile.Close()
}
