package main

import (
	"fmt"
	"os"

	"github.com/tealeg/xlsx/v3"

	"js.comp.dispatching/src/config"
	"js.comp.dispatching/src/models"
	"js.comp.dispatching/src/utility"
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
	gansunFileName := ""

	if len(args) > 2 {
		gansunFileName = args[2]
	}

	if cjFileName != "0" {
		utility.CjContain = true
	}

	if gansunFileName != "" && gansunFileName != "0" {
		utility.GansunContain = true
	}

	// Init Config
	utility.ConfigFile = new(config.Config)
	utility.GansunFile = new(config.Gansun)
	config.InitConfig(utility.ConfigFile)
	config.InitGansun(utility.GansunFile)

	j2SheetName, cjSheetName, gansunSheetName, resultFileName, parseType, companyFilter, errStr := utility.GetParameters()
	// j2SheetName := "배차 내역"
	// cjSheetName := "sheet1"
	// gansunSheetName := "sheet1"
	// resultFileName := "result.csv"
	// parseType := "mp"
	// companyFilter := "mp"
	// errStr := ""
	if j2SheetName == "" || cjSheetName == "" || gansunSheetName == "" || resultFileName == "" {
		fmt.Println(models.InputErr)
		return
	}

	if errStr != "" {
		fmt.Println(errStr)
		return
	}

	// Open an existing file and sheets
	j2, err := xlsx.OpenFile("./" + j2FileName)
	if err != nil {
		panic(err)
	}

	cj := new(xlsx.File)
	cjSheet := new(xlsx.Sheet)
	cjData := map[models.SheetComp]models.CompReturn{}
	gansun := new(xlsx.File)
	gansunSheet := new(xlsx.Sheet)
	gansunData := map[models.SheetComp]models.CompReturn{}
	ok := false

	if utility.CjContain {
		cj, err = xlsx.OpenFile("./" + cjFileName)
		if err != nil {
			panic(err)
		}

		cjSheet, ok = cj.Sheet[cjSheetName]
		if !ok {
			fmt.Println(models.CJFileNameErr)
			return
		}

		cjData = utility.ParseCJData(cjSheet, parseType)
	}

	if utility.GansunContain {
		gansun, err = xlsx.OpenFile("./" + gansunFileName)
		if err != nil {
			panic(err)
		}

		gansunSheet, ok = gansun.Sheet[gansunSheetName]
		if !ok {
			fmt.Println(models.GansunFileNameErr)
			return
		}

		gansunData = utility.ParseGansunData(gansunSheet, parseType)
	}

	j2Sheet, ok := j2.Sheet[j2SheetName]
	if !ok {
		fmt.Println(models.J2FileNameErr)
		return
	}

	// Parse Data
	j2Data := utility.ParseJ2Data(j2Sheet, parseType, companyFilter)

	// Compare Data
	comData := utility.CompareData(j2Data, cjData, gansunData)

	// Write Compare Data
	utility.WriteCSVData(resultFileName, comData)

	fmt.Println("Comp Data: ", len(comData))

}
