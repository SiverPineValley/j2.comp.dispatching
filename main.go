package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx/v3"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/transform"

	"js.comp.dispatching/src/config"
	"js.comp.dispatching/src/models"
	"js.comp.dispatching/src/utility"
)

var configFile *config.Config
var cjContain bool = false
var gansunContain bool = false

func main() {
	// Get Args
	args := os.Args[1:]
	if len(args) < 2 {
		fmt.Println(models.ArgsErr)
		return
	}

	j2FileName := args[0]
	cjFileName := args[1]
	gansunFileName := args[2]

	if cjFileName != "0" {
		cjContain = true
	}

	if gansunFileName != "" && gansunFileName != "0" {
		gansunContain = true
	}

	// Init Config
	configFile = new(config.Config)
	config.InitConfig(configFile)

	j2SheetName, cjSheetName, gansunSheetName, resultFileName, parseType := getParameters()
	// j2SheetName := "배차 내역"
	// cjSheetName := "sheet1"
	// resultFileName := "result2.csv"
	// parseType := "2"
	if j2SheetName == "" || cjSheetName == "" || gansunSheetName == "" || resultFileName == "" {
		fmt.Println(models.InputErr)
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

	if cjContain {
		cj, err = xlsx.OpenFile("./" + cjFileName)
		if err != nil {
			panic(err)
		}

		cjSheet, ok = cj.Sheet[cjSheetName]
		if !ok {
			fmt.Println(models.CJFileNameErr)
			return
		}

		cjData = parseCJData(cjSheet, parseType)
	}

	if gansunContain {
		gansun, err = xlsx.OpenFile("./" + gansunFileName)
		if err != nil {
			panic(err)
		}

		gansunSheet, ok = gansun.Sheet[gansunSheetName]
		if !ok {
			fmt.Println(models.GansunFileNameErr)
			return
		}

		gansunData = parseGansunData(gansunSheet, parseType)
	}

	j2Sheet, ok := j2.Sheet[j2SheetName]
	if !ok {
		fmt.Println(models.J2FileNameErr)
		return
	}

	// Parse Data
	j2Data := parseJ2Data(j2Sheet, parseType)

	// Compare Data
	comData := compareData(j2Data, cjData, gansunData)

	// Write Compare Data
	writeCSVData(resultFileName, comData)

	fmt.Println("Comp Data: ", len(comData))

}

func writeXlsxData(comData []models.CompData) {
	// file := xlsx.NewFile()
}

func writeCSVData(resultFileName string, comData []models.CompData) {
	// Comp 파일 생성
	resFile, err := os.Create("./" + resultFileName)
	if err != nil {
		panic(err)
	}

	// Comp csv writer 생성
	w := bufio.NewWriter(resFile)
	resWr := transform.NewWriter(w, korean.EUCKR.NewEncoder())

	// Comp 내용 쓰기
	if cjContain && gansunContain {
		resWr.Write([]byte("NO, 날짜, 차량번호, 출발, 도착, 간선, J2, CJ, Gansun, J2_NO, CJ_NO, Gansun_No, 비고\n"))
	} else if cjContain && !gansunContain {
		resWr.Write([]byte("NO, 날짜, 차량번호, 출발, 도착, 간선, J2, CJ, J2_NO, CJ_NO, 비고\n"))
	} else if !cjContain && gansunContain {
		resWr.Write([]byte("NO, 날짜, 차량번호, 출발, 도착, 간선, J2, Gansun, J2_NO, Gansun_No, 비고\n"))
	}

	for idx, value := range comData {
		gansun := ""
		if value.Gansun {
			gansun = "간선"
		}

		// 자체 No, 날짜, 차량번호, 출발, 도착, 간선
		each := strconv.Itoa(idx) + ", " + value.Date + ", " + value.LicensePlate + ", " + value.Source + ", " + value.Destination + ", " + gansun

		// Ture, False 및 No, 비고
		if cjContain && gansunContain {
			if value.J2 && value.CJ {
				each = each + ",TRUE, TRUE, FALSE, " + getArrayData(value.J2No, " ") + "," + getArrayData(value.CJNo, " ") + "," + "," + getArrayData(value.Reference, ",")
			} else if value.J2 && value.Gansun {
				each = each + ",TRUE, FALSE, TRUE, " + getArrayData(value.J2No, " ") + "," + "," + getArrayData(value.GansunNo, " ") + "," + getArrayData(value.Reference, ",")
			} else if value.J2 && !value.CJ && !value.Gansun {
				each = each + ",TRUE, FALSE, FALSE, " + getArrayData(value.J2No, " ")
			} else if !value.J2 && value.CJ && !value.Gansun {
				each = each + ",FALSE, TRUE, FALSE, , " + getArrayData(value.CJNo, " ")
			} else if !value.J2 && !value.CJ && value.Gansun {
				each = each + ",FALSE, FALSE, TRUE, , , " + getArrayData(value.GansunNo, " ")
			}
		} else if cjContain && !gansunContain {
			if value.J2 && !value.CJ {
				each = each + ",TRUE, FALSE, " + getArrayData(value.J2No, " ")
			} else if !value.J2 && value.CJ {
				each = each + ",FALSE, TRUE, ," + getArrayData(value.CJNo, " ") + "," + getArrayData(value.Reference, ",")
			} else {
				each = each + ",TRUE, TRUE, " + getArrayData(value.J2No, " ") + "," + getArrayData(value.CJNo, " ") + "," + getArrayData(value.Reference, ",")
			}
		} else if !cjContain && gansunContain {
			if value.J2 && !value.Gansun {
				each = each + ",TRUE, FALSE, " + getArrayData(value.J2No, " ")
			} else if !value.J2 && value.Gansun {
				each = each + ",FALSE, TRUE, ," + getArrayData(value.GansunNo, " ") + "," + getArrayData(value.Reference, ",")
			} else {
				each = each + ",TRUE, TRUE, " + getArrayData(value.J2No, " ") + "," + getArrayData(value.GansunNo, " ") + "," + getArrayData(value.Reference, ",")
			}
		}

		each = each + "\n"
		resWr.Write([]byte(each))
		w.Flush()
	}

	resWr.Close()
}

func getArrayData(arr []string, split string) string {
	if len(arr) == 0 {
		return ""
	}

	no := arr[0]

	for idx := 1; idx < len(arr); idx++ {
		no = no + split + arr[idx]
	}
	return no
}

func compareData(j2Data map[models.SheetComp][]string, cjData, gansunData map[models.SheetComp]models.CompReturn) (result []models.CompData) {
	result = make([]models.CompData, 0)
	for key, value := range j2Data {
		cjValue, exists := cjData[key]
		// J2에만 있으면 일단 넘어감
		if !exists {
			continue
		}

		// 둘다 있고 개수 동일하면 둘다 추가 X
		// 개수 다르면 둘다 추가 O
		if len(value) == len(cjValue.Idx) {
			delete(j2Data, key)
			delete(cjData, key)
			continue
		} else {
			result = append(result, models.CompData{
				Date:         key.Date,
				LicensePlate: key.LicensePlate,
				Source:       key.Source,
				Destination:  key.Destination,
				IsGansun:     key.Gansun,
				J2No:         value,
				CJNo:         cjValue.Idx,
				Reference:    cjValue.Reference,
				J2:           true,
				CJ:           true,
				Gansun:       false})
			delete(j2Data, key)
			delete(cjData, key)
			continue
		}
	}

	for key, value := range j2Data {
		gansunValue, exists := gansunData[key]
		// J2에만 있으면 J2만 있다고 추가
		if !exists {
			result = append(result, models.CompData{
				Date:         key.Date,
				LicensePlate: key.LicensePlate,
				Source:       key.Source,
				Destination:  key.Destination,
				IsGansun:     key.Gansun,
				J2No:         value,
				J2:           true,
				CJ:           false,
				Gansun:       false})
			delete(j2Data, key)
			continue
		}

		// 둘다 있고 개수 동일하면 둘다 추가 X
		// 개수 다르면 둘다 추가 O
		if len(value) == len(gansunValue.Idx) {
			delete(j2Data, key)
			delete(gansunData, key)
			continue
		} else {
			result = append(result, models.CompData{
				Date:         key.Date,
				LicensePlate: key.LicensePlate,
				Source:       key.Source,
				Destination:  key.Destination,
				IsGansun:     key.Gansun,
				J2No:         value,
				GansunNo:     gansunValue.Idx,
				Reference:    gansunValue.Reference,
				J2:           true,
				CJ:           false,
				Gansun:       true})
			delete(j2Data, key)
			delete(gansunData, key)
			continue
		}
	}

	// J2에만 있는 녀석
	for key, value := range j2Data {
		result = append(result, models.CompData{
			Date:         key.Date,
			LicensePlate: key.LicensePlate,
			Source:       key.Source,
			Destination:  key.Destination,
			IsGansun:     key.Gansun,
			J2No:         value,
			J2:           true,
			CJ:           false,
			Gansun:       false})
	}

	// CJ에만 있는 녀석
	for key, value := range cjData {
		result = append(result, models.CompData{
			Date:         key.Date,
			LicensePlate: key.LicensePlate,
			Source:       key.Source,
			Destination:  key.Destination,
			IsGansun:     key.Gansun,
			CJNo:         value.Idx,
			Reference:    value.Reference,
			J2:           false,
			CJ:           true,
			Gansun:       false})
	}

	// 간선에만 있는 녀석
	for key, value := range gansunData {
		result = append(result, models.CompData{
			Date:         key.Date,
			LicensePlate: key.LicensePlate,
			Source:       key.Source,
			Destination:  key.Destination,
			IsGansun:     key.Gansun,
			GansunNo:     value.Idx,
			Reference:    value.Reference,
			J2:           false,
			CJ:           false,
			Gansun:       true})
	}

	return
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

func parseJ2Data(j2Sheet *xlsx.Sheet, parseType string) map[models.SheetComp][]string {
	var j2Titles = [...]string{configFile.J2.No, configFile.J2.Date, configFile.J2.LicensePlate, configFile.J2.Route, configFile.J2.Reference, configFile.J2.TargetCompany}
	var startIdx = configFile.J2.StartIdx
	j2TitleIdx := getCellTitle(j2Sheet, j2Titles[:], startIdx)

	noIdx := j2TitleIdx[0]
	dateIdx := j2TitleIdx[1]
	licenceIdx := j2TitleIdx[2]
	routeIdx := j2TitleIdx[3]
	referenceIdx := j2TitleIdx[4]
	targetCompanyIdx := j2TitleIdx[5]

	result := make(map[models.SheetComp][]string)
	for idx := 0 + startIdx + 1; idx < j2Sheet.MaxRow; idx++ {
		// 청구업체
		targetCompanyCell, _ := j2Sheet.Cell(idx, targetCompanyIdx)
		targetCompany := targetCompanyCell.String()

		if value, exists := configFile.Target["filter"]; exists && len(value) > 0 && !utility.Contains(value, targetCompany) {
			continue
		}

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

		// 비고
		referenceCell, _ := j2Sheet.Cell(idx, referenceIdx)
		reference := referenceCell.String()

		// 시프트면 하루 + 1, 토요일이면 + 2
		if strings.Contains(reference, "시프트") {
			if date.Weekday() == time.Saturday {
				date = date.AddDate(0, 0, 2)
			} else {
				date = date.AddDate(0, 0, 1)
			}
		}

		if len(slice) < 2 {
			continue
		}

		sourceLayover := slice[0]
		dest := slice[1]
		isGansun := false
		if strings.Contains(dest, "간선") {
			isGansun = true
		}
		dest = strings.Replace(dest, "간선", "", -1)        // 간선 제거
		dest = strings.Replace(dest, "간선대체", "", -1)      // 간선대체 제거
		sliceLayover := strings.Split(sourceLayover, "/") // Split Layover
		source := sliceLayover[0]
		source = checkLayover(source)

		if no == "" && date.String() == "" && licensePlate == "" && source == "" && dest == "" {
			break
		}

		// 간선, CJ 비교하는지 확인
		if (!gansunContain && isGansun) || (!cjContain && !isGansun) {
			continue
		}

		if checkGetData(source, dest, parseType) {
			each := new(models.SheetComp)
			each.Date = date.String()
			each.LicensePlate = licensePlate
			each.Source = source
			each.Destination = dest
			each.Gansun = isGansun

			if _, exists := result[*each]; !exists {
				result[*each] = make([]string, 0)
			}

			result[*each] = append(result[*each], no)
		}
	}

	return result
}

func parseCJData(cjSheet *xlsx.Sheet, parseType string) map[models.SheetComp]models.CompReturn {
	var cjTitles = [...]string{configFile.Cj.No, configFile.Cj.Date, configFile.Cj.LicensePlate, configFile.Cj.Source, configFile.Cj.Destination, configFile.Cj.Reference} // configFile.Cj.LayoverNum, configFile.Cj.CarType}
	cjTitleIdx := getCellTitle(cjSheet, cjTitles[:], 0)
	var startIdx = configFile.Cj.StartIdx

	noIdx := cjTitleIdx[0]
	dateIdx := cjTitleIdx[1]
	licenceIdx := cjTitleIdx[2]
	sourceIdx := cjTitleIdx[3]
	destIdx := cjTitleIdx[4]
	// layoverNumIdx := cjTitleIdx[5]
	// carTypeIdx := cjTitleIdx[6]
	referenceIdx := cjTitleIdx[5]

	result := make(map[models.SheetComp]models.CompReturn)
	for idx := 0 + startIdx; idx < cjSheet.MaxRow; idx++ {
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
		source = strings.Replace(source, " ", "", -1)           // Trim
		source = strings.Replace(source, "이천MP", "이천", -1)      // 이천MP -> 이천
		source = strings.Replace(source, "이천MPHub", "이천MP", -1) // 이천MPHub -> 이천MP
		source = strings.Replace(source, "Sub", "", -1)         // Sub 제거
		source = strings.Replace(source, "Hub", "", -1)         // Hub 제거
		source = strings.Replace(source, "콘솔", "", -1)          // 콘솔 제거
		source = checkLayover(source)

		// 도착
		destCell, _ := cjSheet.Cell(idx, destIdx)
		dest := destCell.String()
		dest = strings.Replace(dest, " ", "", -1)   // Trim
		dest = strings.Replace(dest, "Sub", "", -1) // Sub 제거
		dest = strings.Replace(dest, "Hub", "", -1) // Hub 제거
		dest = strings.Replace(dest, "콘솔", "", -1)  // 콘솔 제거

		// 경유 횟수
		// layoverNumCell, _ := cjSheet.Cell(idx, layoverNumIdx)
		// layoverNum := layoverNumCell.String()

		// 차량 구분
		// carTypeCell, _ := cjSheet.Cell(idx, carTypeIdx)
		// carType := carTypeCell.String()

		// 비고
		referenceCell, _ := cjSheet.Cell(idx, referenceIdx)
		reference := referenceCell.String()
		if strings.Contains(reference, ",") {
			reference = "\"" + reference + "\""
		}

		if no == "합계" || no == "" && date.String() == "" && licensePlate == "" && source == "" && dest == "" {
			break
		}

		if checkGetData(source, dest, parseType) {
			each := new(models.SheetComp)
			each.Date = date.String()
			each.LicensePlate = licensePlate
			each.Source = source
			each.Destination = dest
			each.Gansun = false
			// layoverNumInt, err := strconv.Atoi(layoverNum)
			// if err != nil {
			// 	layoverNumInt = 0
			// }

			if _, exists := result[*each]; !exists {
				value := new(models.CompReturn)
				result[*each] = *value
			}

			value := result[*each]
			value.Idx = append(value.Idx, no)
			value.Reference = append(value.Reference, reference)
			result[*each] = value
		}
	}

	return result

}

func parseGansunData(gansunSheet *xlsx.Sheet, parseType string) map[models.SheetComp]models.CompReturn {
	var gansunTitles = [...]string{configFile.Gansun.No, configFile.Gansun.Date, configFile.Gansun.LicensePlate, configFile.Gansun.Source, configFile.Gansun.Destination, configFile.Gansun.Reference}
	gansunTitleIdx := getCellTitle(gansunSheet, gansunTitles[:], 0)
	var startIdx = configFile.Gansun.StartIdx

	noIdx := gansunTitleIdx[0]
	dateIdx := gansunTitleIdx[1]
	licenceIdx := gansunTitleIdx[2]
	sourceIdx := gansunTitleIdx[3]
	destIdx := gansunTitleIdx[4]
	referenceIdx := gansunTitleIdx[5]

	result := make(map[models.SheetComp]models.CompReturn)
	for idx := 0 + startIdx; idx < gansunSheet.MaxRow; idx++ {
		// No
		noCell, _ := gansunSheet.Cell(idx, noIdx)
		no := noCell.String()

		// 날짜
		dateCell, _ := gansunSheet.Cell(idx, dateIdx)
		date, _ := dateCell.GetTime(false)

		// 차량번호
		licenseCell, _ := gansunSheet.Cell(idx, licenceIdx)
		licensePlate := licenseCell.String()

		// 출발
		sourceCell, _ := gansunSheet.Cell(idx, sourceIdx)
		source := sourceCell.String()
		source = strings.Replace(source, " ", "", -1)   // Trim
		source = strings.Replace(source, "Sub", "", -1) // Sub 제거
		source = strings.Replace(source, "Hub", "", -1) // Hub 제거
		source = strings.Replace(source, "콘솔", "", -1)  // 콘솔 제거
		source = checkLayover(source)

		// 도착
		destCell, _ := gansunSheet.Cell(idx, destIdx)
		dest := destCell.String()
		dest = strings.Replace(dest, " ", "", -1)           // Trim
		dest = strings.Replace(dest, "이천MP", "이천", -1)      // 이천MP -> 이천
		dest = strings.Replace(dest, "이천MPHub", "이천MP", -1) // 이천MPHub -> 이천MP
		dest = strings.Replace(dest, "Sub", "", -1)         // Sub 제거
		dest = strings.Replace(dest, "Hub", "", -1)         // Hub 제거
		dest = strings.Replace(dest, "콘솔", "", -1)          // 콘솔 제거

		// 비고
		referenceCell, _ := gansunSheet.Cell(idx, referenceIdx)
		reference := referenceCell.String()
		if strings.Contains(reference, ",") {
			reference = "\"" + reference + "\""
		}

		if no == "합계" || no == "" && date.String() == "" && licensePlate == "" && source == "" && dest == "" {
			break
		}

		if checkGetData(source, dest, parseType) {
			each := new(models.SheetComp)
			each.Date = date.String()
			each.LicensePlate = licensePlate
			each.Source = dest
			each.Destination = source
			each.Gansun = true

			if _, exists := result[*each]; !exists {
				value := new(models.CompReturn)
				result[*each] = *value
			}

			value := result[*each]
			value.Idx = append(value.Idx, no)
			value.Reference = append(value.Reference, reference)
			result[*each] = value
		}
	}

	return result

}

func getParameters() (j2SheetName, cjSheetName, gansunSheetName, resultFileName, resultType string) {
	fmt.Print("J2 파일 시트명(Default: sheet1): ")
	j2Buf := bufio.NewScanner(os.Stdin)
	j2Buf.Scan()
	j2SheetName = j2Buf.Text()

	if !cjContain {
		fmt.Print("CJ 파일 시트명(Default: sheet1): ")
		cjBuf := bufio.NewScanner(os.Stdin)
		cjBuf.Scan()
		cjSheetName = cjBuf.Text()
	}

	if !gansunContain {
		fmt.Print("간선 파일 시트명(Default: sheet1): ")
		gansunBuf := bufio.NewScanner(os.Stdin)
		gansunBuf.Scan()
		gansunSheetName = gansunBuf.Text()
	}

	fmt.Print("결과 파일명(Default: result.csv): ")
	resBuf := bufio.NewScanner(os.Stdin)
	resBuf.Scan()
	resultFileName = resBuf.Text()

	fmt.Print("파싱 타켓(TOML파일 target): ")
	fmt.Scanln(&resultType)

	if j2SheetName == "" {
		j2SheetName = "sheet1"
	}

	if cjSheetName == "" {
		cjSheetName = "sheet1"
	}

	if gansunSheetName == "" {
		gansunSheetName = "sheet1"
	}

	if resultFileName == "" {
		resultFileName = "result.csv"
	}

	if _, exists := configFile.Target[resultType]; !exists {
		resultType = "1"
	}

	return
}

// 가져오는 데이터가 맞는지 조건 확인
func checkGetData(source, dest, parseType string) bool {
	for _, target := range configFile.Target[parseType] {
		if strings.Contains(source, target) || strings.Contains(dest, target) {
			return true
		}
	}

	return false
}

// Layover Check
func checkLayover(key string) string {
	if value, exists := configFile.Direct[key]; exists {
		return value
	}
	return key
}
