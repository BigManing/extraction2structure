package main

import (
	"bytes"
	"fmt"
	"strings"
	"unicode/utf8"

	"github.com/tealeg/xlsx"
)
/**
BigManing
lingshui2008@qq.com
根据excle表格 生成结构体
*/
func main() {
	creatStruct()

}

var bufNotes bytes.Buffer
var bufStruct bytes.Buffer

var sheetName string

func creatStruct() {
	excelFileName := "bean.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}
	for _, sheet := range xlFile.Sheets {
		// 需要指定的sheet
		sheetName = sheet.Name
		// fmt.Printf("Sheet Name: %s\n", sheet.Name)
		//  重置 buf
		bufNotes.Reset()
		bufStruct.Reset()

		//   计算 字段名字 的最大长度   只计算 第二行 第一列的所有的数据
		maxNoteLength, maxStructLength := calculateMaxLength(sheet)

		println()
		for rowIndex, row := range sheet.Rows {
			//  第一行不要
			if rowIndex == 0 {
				continue
			}
			//  如果是空的一行  就退出
			if row.Cells[0].String() == "" {
				break
			}
			//  从第二行开始
			for cellIndex, cell := range row.Cells {
				//  只需要前两列
				if cellIndex > 1 {
					break
				}
				// 如果是空字符串直接break
				text := cell.String()

				addNotes(rowIndex, cellIndex, maxNoteLength, text)

				addStruct(rowIndex, cellIndex, text, sheet.Name, maxStructLength)
			}
			bufNotes.WriteString("\n")
			bufStruct.WriteString("\n")
		}
		bufStruct.WriteString("}")

		// 最后 打印
		println(bufNotes.String())
		println(bufStruct.String())
		println()
	}
}

type AA struct {
	Ok int `json:"ok"`
}

/**
 * 添加结构体
 */
func addStruct(rowIndex int, cellIndex int, text string, sheetName string, maxStructLength int) {

	// 只关注 第二列的数据
	if cellIndex == 1 {
		if rowIndex == 1 {
			bufStruct.WriteString(fmt.Sprintf("type %s struct {\n", sheetName))
		}
		//  根据字段长度  合理设置  空格
		formatStr := "%s"
		count := maxStructLength + 2 - utf8.RuneCountInString(text)
		for index := 0; index < count; index++ {
			//  一个汉字 两个字节
			formatStr += " "
		}
		// 最后跟上类型   json tag
		formatStr += "string  `json:\"%s\"`"
		bufStruct.WriteString(fmt.Sprintf(formatStr, strFirstToUpper(text), text))
	}
}

/**
 * 添加注释
 */
func addNotes(rowIndex int, cellIndex int, maxLength int, text string) {

	//  只有第一列 才需要添加 ‘//’  第二列不需要添加
	if cellIndex == 0 {

		//  第一行的时候需要 添加类名
		if rowIndex == 1 {
			bufNotes.WriteString(fmt.Sprintf("//-----------%s-----------\n", sheetName))
		}

		count := maxLength + 2 - utf8.RuneCountInString(text)
		for index := 0; index < count; index++ {
			//  一个汉字 两个字节
			text += "  "
		}
		text = "// " + text
	}
	bufNotes.WriteString(text)
}

/**
 * 字符串首字母转化为大写 teacher_no -> TeacherNo
 */
func strFirstToUpper(str string) string {
	temp := strings.Split(str, "_")
	var upperStr string
	for y := 0; y < len(temp); y++ {
		vv := []rune(temp[y])
		for i := 0; i < len(vv); i++ {
			if i == 0 {
				vv[i] -= 32
				upperStr += string(vv[i])
			} else {
				upperStr += string(vv[i])
			}
		}
	}
	return upperStr
}

/**
计算所有字符 最大的那个长度
*/
func calculateMaxLength(sheet *xlsx.Sheet) (int, int) {
	maxNoteLength := 0
	maxStructLength := 0

	for rowIndex, row := range sheet.Rows {
		//  第一行不要
		if rowIndex == 0 {
			continue
		}
		for cellIndex, cell := range row.Cells {
			text := cell.String()
			switch cellIndex {
			// 注释的最大长度(第一列)
			case 0:
				if utf8.RuneCountInString(text) > maxNoteLength {
					maxNoteLength = utf8.RuneCountInString(text)
				}
			// 结构体字段  的最大长度(第二列)
			case 1:
				if utf8.RuneCountInString(text) > maxStructLength {
					maxStructLength = utf8.RuneCountInString(text)
				}
			default:
				break
			}
		}
		// fmt.Println(fmt.Sprintf("---%s---%v-----",text,maxLength))
	}
	return maxNoteLength, maxStructLength
}

