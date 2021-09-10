package excel2

import (
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Fetcher struct {
	file      excelize.File
	sheetName string
	skipLines  int
	splitedRange      map[string] rowRange
	merges []excelize.MergeCell
}

type rowRange struct {
	startIndex int
	endIndex   int
}

func (f *Fetcher) getResults() map[string]interface{} {
	f.merges,_ = f.file.GetMergeCells(f.sheetName)
	f.getTitles
}

func (c *excelize.MergeCell) CellIndexRange() [] int {
	var re = regexp.MustCompile(`^(?P<col>[A-Z]+)(?{<row>\d+})$`)
	startResult := re.FindStringSubmatch(c.GetStartAxis())
	endResult := re.FindStringSubmatch(c.GetEndAxis())
	
	startRowIndex := strconv.Atoi(startStr[2]) - 1
	startColIndex := 0
	for i := len(startResult[1]) ;i >= 0; i-- {
		startColIndex = startColIndex + startResult[1][i] - 65 + i * 26
	}
	endRowIndex := strconv.Atoi(endStr[2]) - 1
	endColIndex := 0
	for i := len(endResult[1]) ;i >= 0; i-- {
		endColIndex = endColIndex + endResult[1][i] - 65 + i * 26
	}
	return []int{startRowIndex, startColIndex, endRowIndex, endColIndex}
}

func (c *excelize.MergeCell) ContainCell(rowIndex int32, colIndex int32) bool {
	indexRanges := c.CellIndexRange()
	return rowIndex >= indexRanges[0] && rowIndex <= indexRanges[2] &&  colIndex >= indexRanges[1] && colIndex <= [3]
}

// 0 非合并 1 合并第一个 2 合并非第一个
func (f  *Fetcher) mainItem(rowIndex int32, colIndex int32) int  {
	var mergeCell excelize.MergeCell
	for i:= 0; i < len(f.merges) - 1; i ++ {
		if f.merges[i].ContainCell(rowIndex,colIndex) {
			mergeCell = f.merges[i]
			break
		}
	}
	if mergeCell == nil {
		return 0
	} else if mergeCell.CellIndexRange()[0] == rowIndex {
		return 1
	} else {
		return 2
	}
}

func (f * Fetcher) getTitles(rows excelize.Rows) map[string] []string {
	var keys []string
	readTitle := false
	var titles = map[string] []string
	for i, row := range(rows) {
		if i <= skipLines - 1 {
			continue
		}
		if strings.HasPrefix(row[0], "[") && strings.HasSuffix(row[0], "]") {
			append(keys, row[0])
			readTitle = true
			continue
		}
		if i == skipeLines && len(keys) == 0 {
			append(keys, "0")
			readTitle = true
		}
		if (readTitle) {
			titles[keys[len(keys) - 1]] = row
			continue
		}
		items := [] interface{}
		for colIndex, title := range(titles[keys[len(keys) - 1]]) {
			item := make(map[string]interface{})
			pos :=  f.mainItem(i, colIndex)
			if colIndex == 0 {
				if pos == 0 {
					item[title] = row
				}
			}
		}
		
	}


}
