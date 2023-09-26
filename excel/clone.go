package excel

import (
	"excel_kit/utils"
	"excel_kit/validator"
	"fmt"
	"github.com/charmbracelet/bubbles/key"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/fzdwx/infinite"
	"github.com/fzdwx/infinite/components"
	"github.com/fzdwx/infinite/components/selection/multiselect"
	"github.com/fzdwx/infinite/components/selection/singleselect"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
)

func CombaineSheetData() {
	// choose the file
	path := chooseFile()
	file, err := excelize.OpenFile(path, excelize.Options{})
	if err != nil {
		log.Printf("打开excel失败")
		fmt.Printf("异常：%+v\n", err)
		return
	}
	defer func() {
		if err := file.Close(); err != nil {
			panic(err)
		}
	}()
	// opt table.
	optTable(file)
}

func chooseFile() string {
	dir, err := os.Getwd()
	if err != nil {
		fmt.Println("无法获取当前目录：", err)
		return ""
	}

	var files []string

	err = filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if !info.IsDir() && (strings.HasSuffix(info.Name(), ".xlsx") || strings.HasSuffix(info.Name(), ".xls")) {
			files = append(files, path)
		}
		return nil
	})

	if err != nil {
		fmt.Println("遍历目录时发生错误：", err)
		return ""
	}
	idx, err := configSingleselect(files, "请选择需要确认的excel文件:[enter:确认]")
	if err != nil {
		return ""
	}

	return files[idx]
}

func chooseSheet(file *excelize.File) ([]string, []int, error) {
	list := file.GetSheetList()

	display, err := configMultiselect(list, "请选择需要匹配的工作薄名称[空格：选中；enter:确认]：")
	if err != nil {
		return nil, nil, err
	}
	return list, display, err
}

// configMultiselect encapsulated the component of which is multiselect.
func configMultiselect(list []string, msg string) ([]int, error) {
	keyMap := components.DefaultMultiKeyMap()
	keyMap.Confirm = key.NewBinding(
		key.WithKeys("enter"),
	)
	keyMap.Choice = key.NewBinding(
		key.WithKeys(tea.KeySpace.String()),
	)
	selected := infinite.NewMultiSelect(list,
		multiselect.WithKeyMap(keyMap),
		multiselect.WithHintSymbol("√"),
		multiselect.WithUnHintSymbol("x"),
		multiselect.WithValidator(validator.Min(2, "需要匹配的工作薄至少需要选中%v个以上")),
		multiselect.WithPageSize(10),
		multiselect.WithDisableFilter())

	display, err := selected.Display(msg)
	return display, err
}

func configSingleselect(choices []string, pop string) (int, error) {
	keyMap := singleselect.DefaultSingleKeyMap()
	keyMap.Choice = key.NewBinding(
		key.WithKeys("enter"),
	)
	keyMap.Confirm = key.NewBinding(
		key.WithKeys("enter"),
	)

	display, err := infinite.NewSingleSelect(choices,
		singleselect.WithDisableFilter(),
		singleselect.WithKeyBinding(keyMap),

		singleselect.WithPageSize(5),
	).Display(pop)

	if err != nil {
		return 0, err
	}
	return display, err
}

// chooseColCell choose the special cell from all of which are some chosed sheets.
func chooseColCell(file *excelize.File, sheets []string, arr []int) (map[string][]string, []map[string][]string) {
	firstData := make(map[string][]string)
	allMatchslice := make([]map[string][]string, 0)
	// arr 为选中的多个工作薄
	for idx, val := range arr {
		// 获取当前工作薄
		cur := sheets[val]
		curmap := make(map[string][]string)
		// 获取当前工作薄上所有的单元格
		rows, err := file.GetRows(cur)
		// 获取选中的列单元
		cell, err := configSingleselect(rows[0], "请选择工作薄:"+cur+"需要匹配的列名[enter:确认]")
		for i, row := range rows {
			if idx == 0 {
				if i == 0 {
					firstData["title"] = row
				} else {
					firstData[row[cell]] = row
				}
			} else {

				newSlice := make([]string, len(row)-1)
				copy(newSlice, row[:cell])
				copy(newSlice[cell:], row[cell+1:])
				if i == 0 {
					curmap["title"] = newSlice
				} else {
					curmap[row[cell]] = newSlice
				}
			}
		}
		if idx > 0 {
			allMatchslice = append(allMatchslice, curmap)
		}
		if err != nil {
			panic(err)
		}
	}

	return firstData, allMatchslice
}

// optTable opearte the table.
func optTable(file *excelize.File) {
	sheets, arr, err := chooseSheet(file)
	firstData, allMatchDatas := chooseColCell(file, sheets, arr)
	reg := regexp.MustCompile(`[\p{Han}]+`)
	space := regexp.MustCompile(`\s+`)
	// allMatchDatas 切片map
	for _, data := range allMatchDatas {
		// firstData map
		for k := range firstData {
			if k == "title" {
				firstData["title"] = append(firstData["title"], data["title"]...)

			} else {
				for ke, va := range data {
					if ke == "title" {
						continue
					}
					matches := strings.Join(reg.FindAllString(ke, -1), "")
					matches = space.ReplaceAllString(matches, "")
					curkey := strings.Join(reg.FindAllString(k, -1), "")
					curkey = space.ReplaceAllString(curkey, "")
					if strings.Contains(curkey, matches) {
						firstData[k] = append(firstData[k], va...)
						delete(data, ke)
					}
				}
			}
		}
	}
	// combaine the unique elements to firstData.
	fmt.Println("剩余元素：", allMatchDatas)
	extendColLen := len(allMatchDatas[0]["title"])
	delete(allMatchDatas[0], "title")
	if len(allMatchDatas) > 0 && !utils.ContainsKey(allMatchDatas[0], "title") {
		cols := len(firstData["title"])
		for _, data := range allMatchDatas {
			for k, v := range data {
				sli := []string{k}
				for i := 1; i < (cols - extendColLen); i++ {
					sli = append(sli, "")
				}
				newVal := append(sli, v...)
				firstData[k] = append(firstData[k], newVal...)
			}
		}
	}
	fmt.Println("firstData:", firstData)
	list := make([]string, 0, len(firstData))
	for k := range firstData {
		list = append(list, k)
	}
	sort.Strings(list)

	// 根据组合的数据生成新的excel文件
	nf := excelize.NewFile()
	defer func() {
		if err := nf.Close(); err != nil {
			panic(err)
		}
	}()
	// 创建一个工作表
	name := "数据整合"
	ns, err := nf.NewSheet(name)

	// 写入sheet数据
	for idx, d := range list {
		value := firstData[d]
		cell, err := excelize.CoordinatesToCellName(1, idx+1)
		err = nf.SetSheetRow(name, cell, &value)
		if err != nil {
			panic(err)
		}
		configExcelStyle(nf, name, value)
	}
	// 设置工作簿的默认工作表
	nf.SetActiveSheet(ns)
	err = nf.SaveAs("数据整合后表格.xlsx")
	if err != nil {
		panic(err)
	}
}

func configExcelStyle(f *excelize.File, sheet string, data []string) int {

	name, err := excelize.ColumnNumberToName(len(data))
	err = f.SetColWidth(sheet, "A", name, 11.68)
	style, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size:   11,
			Family: "微软雅黑",
			Color:  "000000",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
		},
	})

	if err != nil {
		panic(err)
	}
	return style
}
