package main

import (
	"fmt"
	_ "image/png"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// func writeFile() {
// 	f := excelize.NewFile()

// 	file, err := ioutil.ReadFile("aaa.png")
// 	if err != nil {
// 		fmt.Println(err)
// 	}
// 	if err := f.AddPictureFromBytes("Sheet1", "B2", "", "Excel Logo", ".png", file); err != nil {
// 		fmt.Println("eee")
// 		fmt.Println(err)
// 	}
// 	if err := f.SaveAs("Book1.xlsx"); err != nil {
// 		fmt.Println(err)
// 	}
// }

func readFile() {
	f, err := excelize.OpenFile("e1.xlsx")
	if err != nil {
		fmt.Println("sss")
		fmt.Println(err)
		return
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	for a, row := range rows {
		for b, _ := range row {
			cellKey := string(b+65) + strconv.Itoa(a)
			v, _ := f.GetCellValue("Sheet1", cellKey)
			fmt.Printf("%v:%v;", cellKey, v)
		}
		fmt.Println()
	}

	// merges, error := f.GetMergeCells("Sheet1")
	// if error != nil {
	// 	fmt.Println(err)
	// }
	// titles := cells[0]

	// for i, cell := range cells {
	// 	if i d
	// 	fmt.Println(i, cell[0], cell[1])
	// }

	// file, raw, err := f.GetPicture("Sheet1", "N4")
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }
	// fmt.Println(file)
	// // fmt.Println(raw)
	// if err := ioutil.WriteFile(file, raw, 0644); err != nil {
	// 	fmt.Println(err)
	// }
}
func main() {
	// writeFile()
	readFile()
}
