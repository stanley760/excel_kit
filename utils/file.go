package utils

import (
	"fmt"
	"os"
	"strings"
	"time"
)

func DeleteFile(filePath string) bool {
	// 判断文件是否存在
	if !ExistFile(filePath) {
		return false
	}

	// 删除文件
	err := os.Remove(filePath)
	if err != nil {
		fmt.Println("删除文件出错:", err)
		return false
	}
	return true
}

func ExistFile(fileName string) bool {
	_, err := os.Stat(fileName)
	if err != nil {
		if os.IsNotExist(err) {
			fmt.Println("文件不存在")
		} else {
			fmt.Println("其他错误:", err)
		}
		return false
	}
	return true
}

func NewFileNameWithTimestramp(name string) string {
	split := strings.Split(name, ".")
	now := time.Now().Format("2006-01-02_15-04-05")
	return fmt.Sprintf("%s-%s.%s", split[0], now, split[1])
}
