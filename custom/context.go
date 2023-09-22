package custom

import (
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
	"log"
	"os/exec"
)

// ApplicationContext struct the layout with some widget
func ApplicationContext(w fyne.Window) {
	// get the file path
	f1, fv := chooseNewFile(w)
	// get the object of excelize.
	file := getExcelFile(fv)
	// get all sheets from the chosed file.
	sheets := file.GetSheetList()
	// get the choice from the file's sheets.
	f2, chosed := chosedSheets(sheets)
	chooseColumnDiff(chosed)

	// set into the one container.
	ctr := container.NewVBox(f1, f2)
	w.SetContent(ctr)
}

// chooseNewFile choose the file from the file management.
func chooseNewFile(w fyne.Window) (*fyne.Container, string) {
	// the file choose from the new window which is file management.
	fileLable := widget.NewLabel("excel文件路径:")
	defaultFileName := "original.xlsx"
	fileEntry := widget.NewEntry()
	fileEntry.SetPlaceHolder(defaultFileName)
	fileButton := widget.NewButton("打开", func() {
		// todo 改为window api 调用文件管理系统
		exec.Command("explorer")
		nfo := dialog.NewFileOpen(func(closer fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, w)
				return
			}
			if closer == nil {
				log.Println("文件管理窗口已取消")
				return
			} else {
				fileEntry.SetText(closer.URI().Path())
			}
		}, w)
		nfo.SetConfirmText("选中")
		nfo.SetDismissText("取消")
		nfo.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx", ".xlx"}))
		nfo.Show()
	})

	if fileEntry.Text == "" {
		fileEntry.SetText(defaultFileName)
	}
	// 文件选择器
	return container.NewBorder(layout.NewSpacer(), layout.NewSpacer(), fileLable, fileButton, fileEntry), fileEntry.Text
}

// chosedSheets choose two sheets and above which compare with each other.
func chosedSheets(sheets []string) (*fyne.Container, *widget.CheckGroup) {
	comboLable := widget.NewLabel("选择需要对比的工作薄:")

	combo := widget.NewCheckGroup(sheets, func(value []string) {
		log.Println("Select set to", value)
	})
	combo.Horizontal = true
	return container.NewHBox(comboLable, combo), combo
}

// chooseColumnDiff
func chooseColumnDiff(sheets *widget.CheckGroup) {
	sheets.OnChanged = func(strings []string) {
		fmt.Println("chosed sheets:", strings)
	}
}

func getExcelFile(path string) *excelize.File {
	file, err := excelize.OpenFile(path, excelize.Options{})
	if err != nil {
		log.Printf("打开excel失败")
		fmt.Printf("异常：%+v\n", err)
		panic(err)
	}
	defer func() {
		if err := file.Close(); err != nil {
			panic(err)
		}
	}()
	return file
}
