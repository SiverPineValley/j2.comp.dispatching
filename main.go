package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"

	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/transform"

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

	// Parse Data
	j2Data := parseJ2Data(j2Sheet)
	cjData := parseCJData(cjSheet)

	// J2 파일 생성
	j2File, err := os.Create("./j2.csv")
	if err != nil {
		panic(err)
	}

	// J2 csv writer 생성
	j2Wr := transform.NewWriter(bufio.NewWriter(j2File), korean.EUCKR.NewEncoder())

	// J2 내용 쓰기
	j2Wr.Write([]byte("날짜, 차량번호, 출발, 도착, NO\n"))
	for key, value := range j2Data {
		each := key.Date + ", " + key.LicensePlate + ", " + key.Source + ", " + key.Destination
		no := value[0]

		for idx := 1; idx < len(value); idx++ {
			no = no + " " + value[idx]
		}

		each = each + ", " + no + "\n"
		j2Wr.Write([]byte(each))
	}

	j2Wr.Close()

	// CJ 파일 생성
	cjFile, err := os.Create("./cj.csv")
	if err != nil {
		panic(err)
	}

	// CJ csv writer 생성
	cjWr := transform.NewWriter(bufio.NewWriter(cjFile), korean.EUCKR.NewEncoder())

	// CJ 내용 쓰기
	cjWr.Write([]byte("날짜, 차량번호, 출발, 도착, NO\n"))
	for key, value := range j2Data {
		each := key.Date + ", " + key.LicensePlate + ", " + key.Source + ", " + key.Destination
		no := value[0]

		for idx := 1; idx < len(value); idx++ {
			no = no + " " + value[idx]
		}

		each = each + ", " + no + "\n"
		cjWr.Write([]byte(each))
	}

	cjWr.Close()

	fmt.Println(len(j2Data), "<- j2 길이, cj 길이 ->", len(cjData))
}

func getCellTitle(sh *xlsx.Sheet, title []string, startIdx int) []int {
	length := len(title)
	res := make([]int, length)

	for key, value := range title {
		for idx := 0; idx < sh.MaxCol; idx++ {
			c, _ := sh.Cell(0+startIdx, idx)
			if value == c.String() {
				res[key] = idx
				continue
			}
		}
	}

	return res
}

func parseJ2Data(j2Sheet *xlsx.Sheet) map[models.SheetComp][]string {
	var j2Titles = [...]string{"no", "날 짜", "차량 번호", "운행 노선"}
	var startIdx = 5
	j2TitleIdx := getCellTitle(j2Sheet, j2Titles[:], startIdx)

	noIdx := j2TitleIdx[0]
	dateIdx := j2TitleIdx[1]
	licenceIdx := j2TitleIdx[2]
	routeIdx := j2TitleIdx[3]

	result := make(map[models.SheetComp][]string)
	for idx := 0 + startIdx + 1; idx < j2Sheet.MaxRow; idx++ {
		// No
		noCell, _ := j2Sheet.Cell(idx, noIdx)
		no := noCell.String()

		// 날짜
		dateCell, _ := j2Sheet.Cell(idx, dateIdx)
		date, _ := dateCell.GetTime(false)

		// 차량번호
		licenseCell, _ := j2Sheet.Cell(idx, licenceIdx)
		licensePlate := licenseCell.String()

		// 운행노선
		routeCell, _ := j2Sheet.Cell(idx, routeIdx)
		route := routeCell.String()
		route = strings.Replace(route, " ", "", -1) // Trim
		slice := strings.Split(route, "-")          // Split

		if len(slice) < 2 {
			continue
		}

		source := slice[0]
		dest := slice[1]

		if no == "" && date.String() == "" && licensePlate == "" && source == "" && dest == "" {
			break
		}

		if strings.Contains(source, "이천MP") || strings.Contains(dest, "이천MP") {
			each := new(models.SheetComp)
			each.Date = date.String()
			each.LicensePlate = licensePlate
			each.Source = source
			each.Destination = dest

			if _, exists := result[*each]; !exists {
				result[*each] = make([]string, 0)
			}

			result[*each] = append(result[*each], no)
			// result = append(result, *each)
		}
	}

	return result
}

func parseCJData(cjSheet *xlsx.Sheet) map[models.SheetComp][]string {
	var cjTitles = [...]string{"NO", "출발일자", "차량번호", "출발터미널", "도착터미널"}
	cjTitleIdx := getCellTitle(cjSheet, cjTitles[:], 0)

	noIdx := cjTitleIdx[0]
	dateIdx := cjTitleIdx[1]
	licenceIdx := cjTitleIdx[2]
	sourceIdx := cjTitleIdx[3]
	destIdx := cjTitleIdx[4]

	result := make(map[models.SheetComp][]string)
	for idx := 0 + 1; idx < cjSheet.MaxRow; idx++ {
		// No
		noCell, _ := cjSheet.Cell(idx, noIdx)
		no := noCell.String()

		// 날짜
		dateCell, _ := cjSheet.Cell(idx, dateIdx)
		date, _ := dateCell.GetTime(false)

		// 차량번호
		licenseCell, _ := cjSheet.Cell(idx, licenceIdx)
		licensePlate := licenseCell.String()

		// 출발
		sourceCell, _ := cjSheet.Cell(idx, sourceIdx)
		source := sourceCell.String()
		source = strings.Replace(source, " ", "", -1)   // Trim
		source = strings.Replace(source, "Sub", "", -1) // Sub 제거
		source = strings.Replace(source, "Hub", "", -1) // Hub 제거

		// 도착
		destCell, _ := cjSheet.Cell(idx, destIdx)
		dest := destCell.String()
		dest = strings.Replace(dest, " ", "", -1)   // Trim
		dest = strings.Replace(dest, "Sub", "", -1) // Sub 제거
		dest = strings.Replace(dest, "Hub", "", -1) // Hub 제거

		if no == "합계" || date.String() == "" && licensePlate == "" && source == "" && dest == "" {
			break
		}

		if strings.Contains(source, "이천MP") || strings.Contains(dest, "이천MP") {
			each := new(models.SheetComp)
			each.Date = date.String()
			each.LicensePlate = licensePlate
			each.Source = source
			each.Destination = dest

			if _, exists := result[*each]; !exists {
				result[*each] = make([]string, 0)
			}

			result[*each] = append(result[*each], no)
		}
	}

	return result

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
