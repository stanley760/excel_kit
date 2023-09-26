package utils

import (
	"os"
	"os/exec"
	"runtime"
)

func OpenFileManager(path string) error {
	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "darwin": // macOS
		cmd = exec.Command("open", path)
	case "windows": // Windows
		cmd = exec.Command("explorer", path)
	default: // Linux 和其他 Unix 系统
		cmd = exec.Command("xdg-open", path)
	}

	cmd.Stderr = os.Stderr
	return cmd.Run()
}
