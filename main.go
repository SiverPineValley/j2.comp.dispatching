package main

import (
	"fmt"
	"os"

	"github.com/tealeg/xlsx"
)

func main() {
	args := os.Args[1:]
	if len(args) < 2 {
		fmt.Println("./main.exe (J2파일) (CJ파일) 로 입력해주세요.")
	}

	j2 := args[0]
	cj := args[1]

	// open an existing file
	wb, err := xlsx.OpenFile("../samplefile.xlsx")
	if err != nil {
		panic(err)
	}
	// wb now contains a reference to the workbook
	// show all the sheets in the workbook
	fmt.Println("Sheets in this file:")
	for i, sh := range wb.Sheets {
		fmt.Println(i, sh.Name)
	}
	fmt.Println("----")
}
