package excel

import (
	"excel_kit/validator"
	"fmt"
	"github.com/charmbracelet/bubbles/key"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/fzdwx/infinite"
	"github.com/fzdwx/infinite/components"
	"github.com/fzdwx/infinite/components/input/text"
	"github.com/fzdwx/infinite/components/selection/multiselect"
	"github.com/fzdwx/infinite/components/selection/singleselect"
	"github.com/fzdwx/infinite/theme"
	"github.com/xuri/excelize/v2"
	"log"
	"regexp"
	"sort"
	"strings"
)

func CombaineSheetData() {
	pathName := infinite.NewText(
		text.WithPrompt("请输入文档路径(可选,默认为当前文件夹下的同名excel文件):"),
		text.WithPromptStyle(theme.DefaultTheme.PromptStyle),
		text.WithDefaultValue("orignal.xlsx"),
	)
	path, err := pathName.Display()
	if err != nil {
		panic(err)
	}
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

func chooseColCell(file *excelize.File, sheets []string, arr []int) (map[string][]string, []map[string][]string) {
	firstData := make(map[string][]string)
	allMatchslice := make([]map[string][]string, 0)
	// arr 为选中的多个工作薄
	for idx := range arr {
		// 获取当前工作薄
		cur := sheets[idx]
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

func optTable(file *excelize.File) {
	sheets, arr, err := chooseSheet(file)
	firstData, allMatchDatas := chooseColCell(file, sheets, arr)
	for _, data := range allMatchDatas {
		for k := range firstData {
			if k == "title" {
				firstData["title"] = append(firstData["title"], data["title"]...)
				continue
			}
			for ke, va := range data {
				reg := regexp.MustCompile(`[\p{Han}]+`)
				matches := strings.Join(reg.FindAllString(ke, -1), "")
				if strings.Contains(k, matches) && ke != "title" {
					firstData[k] = append(firstData[k], va...)
				}
			}
		}
	}
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
	name := "数据整合工作薄"
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
	if err != nil {
		panic(err)
	}
	// 设置工作簿的默认工作表
	nf.SetActiveSheet(ns)
	if err := nf.SaveAs("数据整合后表格.xlsx"); err != nil {
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
