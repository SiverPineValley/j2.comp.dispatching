package main

import (
	"fmt"
	"os"

	"github.com/tealeg/xlsx/v3"
	"js.comp.dispatching/src/models"
)

func main() {
	// Get Args
	args := os.Args[1:]
	if len(args) < 2 {
		fmt.Println(models.ArgsErr)
		return
	}

	j2FileName := args[0]
	cjFileName := args[1]

	j2SheetName, cjSheetName := getSheetName()

	// Open an existing file and sheets
	j2, err := xlsx.OpenFile("./" + j2FileName)
	if err != nil {
		panic(err)
	}

	cj, err := xlsx.OpenFile("./" + cjFileName)
	if err != nil {
		panic(err)
	}

	j2Sheet, ok := j2.Sheet[j2SheetName]
	if !ok {
		fmt.Println(models.J2FileNameErr)
		return
	}

	cjSheet, ok := cj.Sheet[cjSheetName]
	if !ok {
		fmt.Println(models.CJFileNameErr)
		return
	}

	// J2 Parse Data
	j2Date := "날 짜"
	j2LicensePlate := "차량 번호"
	j2Route := "운행 노선"

	j2Title, err := j2Sheet.Row(0)
	if err != nil {
		fmt.Println(models.J2InvalidFileErr)
	}

	// CJ Parse Data
	cjDate := "날 짜"
	cjLicensePlate := "차량 번호"
	cjDeparture := "출발터미널"
	cjDestination := "도착터미널"
	cjTitle, err := cjSheet.Row(0)
	if err != nil {
		fmt.Println(models.CJInvalidFileErr)
	}
}

func cellVisitor(c *xlsx.Cell) error {
	value, err := c.FormattedValue()
	if err != nil {
		fmt.Println(err.Error())
	} else {
		fmt.Println("Cell value:", value)
	}
	return err
}

func rowVisitor(r *xlsx.Row) error {
	return r.ForEachCell(cellVisitor)
}

func getSheetName() (j2SheetName, cjSheetName string) {
	fmt.Print("J2 파일 시트명(Default: sheet1): ")
	fmt.Scanln(&j2SheetName)

	fmt.Print("CJ 파일 시트명(Default: sheet1): ")
	fmt.Scanln(&cjSheetName)

	if j2SheetName == "" {
		j2SheetName = "sheet1"
	}

	if cjSheetName == "" {
		cjSheetName = "sheet1"
	}

	return
}
