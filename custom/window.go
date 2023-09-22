package custom

import (
	"os"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"github.com/flopp/go-findfont"
)

func init() {
	fontPaths := findfont.List()
	for _, fontPath := range fontPaths {
		// 黑体:simhei.ttf
		if strings.Contains(fontPath, "simhei.ttf") {
			err := os.Setenv("FYNE_FONT", fontPath)
			if err != nil {
				panic("加载字体失败")
			}
			break
		}
	}
}

func NewWindows(title string, width, length float32) {
	a := app.New()

	w := a.NewWindow(title)
	w.Resize(fyne.NewSize(width, length))
	ApplicationContext(w)
	w.CenterOnScreen()
	w.ShowAndRun()
	// unsets a customed font
	err := os.Unsetenv("FYNE_FONT")
	if err != nil {
		panic(err)
	}
}
