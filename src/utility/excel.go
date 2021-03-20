package utility

import (
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx/v3"
	"js.comp.dispatching/src/config"
	"js.comp.dispatching/src/models"
)

var ConfigFile *config.Config
var GansunFile *config.Gansun

func CompareData(j2Data, cjData, gansunData map[models.SheetComp]models.CompReturn) (result []models.CompData) {
	result = make([]models.CompData, 0)
	for key, value := range j2Data {
		cjValue, exists := cjData[key]
		// J2에만 있으면 일단 넘어감
		if !exists {
			continue
		}

		// 둘다 있고 개수 동일하면 둘다 추가 X
		// 개수 다르면 둘다 추가 O
		if len(value.Idx) == len(cjValue.Idx) {
			isSame, _, _ := CheckTotalFee(value.TotalFee, cjValue.TotalFee)
			if value.Stage == 2 && isSame {
				value.Stage = 3
			}

			for idx, _ := range value.TotalFee {
				result = append(result, models.CompData{
					Date:             key.Date,
					LicensePlate:     key.LicensePlate,
					Source:           key.Source,
					Destination:      key.Destination,
					IsGansun:         false,
					IsGansunOneway:   false,
					J2No:             value.Idx,
					CJNo:             cjValue.Idx,
					J2Reference:      value.Reference,
					CJReference:      cjValue.Reference,
					DetourFeeType:    cjValue.DetourFeeType,
					DetourFee:        cjValue.DetourFee,
					DetourFeeType3:   cjValue.DetourFeeType3,
					DetourFee3:       cjValue.DetourFee3,
					MultiTourPercent: cjValue.MultiTourPercent,
					J2:               true,
					CJ:               true,
					Gansun:           false,
					Stage:            value.Stage,
					FirstTotalFee:    value.TotalFee[idx],
					SecondTotalFee:   cjValue.TotalFee[idx],
				})
			}

			delete(j2Data, key)
			delete(cjData, key)
			continue
		} else {
			result = append(result, models.CompData{
				Date:             key.Date,
				LicensePlate:     key.LicensePlate,
				Source:           key.Source,
				Destination:      key.Destination,
				IsGansun:         false,
				IsGansunOneway:   false,
				J2No:             value.Idx,
				CJNo:             cjValue.Idx,
				J2Reference:      value.Reference,
				CJReference:      cjValue.Reference,
				DetourFeeType:    cjValue.DetourFeeType,
				DetourFee:        cjValue.DetourFee,
				DetourFeeType3:   cjValue.DetourFeeType3,
				DetourFee3:       cjValue.DetourFee3,
				MultiTourPercent: cjValue.MultiTourPercent,
				J2:               true,
				CJ:               true,
				Gansun:           false,
				Stage:            0,
			})
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
				Date:           key.Date,
				LicensePlate:   key.LicensePlate,
				Source:         key.Source,
				Destination:    key.Destination,
				IsGansun:       key.Gansun,
				IsGansunOneway: key.GansunOneWay,
				J2No:           value.Idx,
				J2Reference:    value.Reference,
				J2:             true,
				CJ:             false,
				Gansun:         false,
				Stage:          0,
			})
			delete(j2Data, key)
			continue
		}

		// 둘다 있고 개수 동일하면 둘다 추가 X
		// 개수 다르면 둘다 추가 O
		if len(value.Idx) == len(gansunValue.Idx) {
			isSame, _, _ := CheckTotalFee(value.TotalFee, gansunValue.TotalFee)
			if value.Stage == 2 && isSame {
				value.Stage = 3
			}

			for idx, _ := range value.TotalFee {
				result = append(result, models.CompData{
					Date:             key.Date,
					LicensePlate:     key.LicensePlate,
					Source:           key.Source,
					Destination:      key.Destination,
					IsGansun:         key.Gansun,
					IsGansunOneway:   key.GansunOneWay,
					J2No:             value.Idx,
					GansunNo:         gansunValue.Idx,
					J2Reference:      value.Reference,
					CJReference:      gansunValue.Reference,
					DetourFeeType:    gansunValue.DetourFeeType,
					DetourFee:        gansunValue.DetourFee,
					DetourFeeType3:   gansunValue.DetourFeeType3,
					DetourFee3:       gansunValue.DetourFee3,
					MultiTourPercent: gansunValue.MultiTourPercent,
					J2:               true,
					CJ:               false,
					Gansun:           true,
					Stage:            value.Stage,
					FirstTotalFee:    value.TotalFee[idx],
					SecondTotalFee:   gansunValue.TotalFee[idx],
				})
			}

			delete(j2Data, key)
			delete(gansunData, key)
			continue
		} else {
			result = append(result, models.CompData{
				Date:             key.Date,
				LicensePlate:     key.LicensePlate,
				Source:           key.Source,
				Destination:      key.Destination,
				IsGansun:         key.Gansun,
				IsGansunOneway:   key.GansunOneWay,
				J2No:             value.Idx,
				GansunNo:         gansunValue.Idx,
				J2Reference:      value.Reference,
				CJReference:      gansunValue.Reference,
				DetourFeeType:    gansunValue.DetourFeeType,
				DetourFee:        gansunValue.DetourFee,
				DetourFeeType3:   gansunValue.DetourFeeType3,
				DetourFee3:       gansunValue.DetourFee3,
				MultiTourPercent: gansunValue.MultiTourPercent,
				J2:               true,
				CJ:               false,
				Gansun:           true,
				Stage:            0,
			})
			delete(j2Data, key)
			delete(gansunData, key)
			continue
		}
	}

	// J2에만 있는 녀석
	for key, value := range j2Data {
		result = append(result, models.CompData{
			Date:           key.Date,
			LicensePlate:   key.LicensePlate,
			Source:         key.Source,
			Destination:    key.Destination,
			IsGansun:       key.Gansun,
			IsGansunOneway: key.GansunOneWay,
			J2No:           value.Idx,
			J2Reference:    value.Reference,
			J2:             true,
			CJ:             false,
			Gansun:         false,
			Stage:          0,
		})
	}

	// CJ에만 있는 녀석
	for key, value := range cjData {
		result = append(result, models.CompData{
			Date:             key.Date,
			LicensePlate:     key.LicensePlate,
			Source:           key.Source,
			Destination:      key.Destination,
			IsGansun:         false,
			IsGansunOneway:   false,
			CJNo:             value.Idx,
			CJReference:      value.Reference,
			DetourFeeType:    value.DetourFeeType,
			DetourFee:        value.DetourFee,
			DetourFeeType3:   value.DetourFeeType3,
			DetourFee3:       value.DetourFee3,
			MultiTourPercent: value.MultiTourPercent,
			J2:               false,
			CJ:               true,
			Gansun:           false,
			Stage:            0,
		})
	}

	// 간선에만 있는 녀석
	for key, value := range gansunData {
		result = append(result, models.CompData{
			Date:             key.Date,
			LicensePlate:     key.LicensePlate,
			Source:           key.Source,
			Destination:      key.Destination,
			IsGansun:         key.Gansun,
			IsGansunOneway:   key.GansunOneWay,
			GansunNo:         value.Idx,
			CJReference:      value.Reference,
			DetourFeeType:    value.DetourFeeType,
			DetourFee:        value.DetourFee,
			DetourFeeType3:   value.DetourFeeType3,
			DetourFee3:       value.DetourFee3,
			MultiTourPercent: value.MultiTourPercent,
			J2:               false,
			CJ:               false,
			Gansun:           true,
			Stage:            0,
		})
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

func ParseJ2Data(j2Sheet *xlsx.Sheet, parseType, companyFilter string) map[models.SheetComp]models.CompReturn {
	var j2Titles = [...]string{
		ConfigFile.J2.No,
		ConfigFile.J2.Date,
		ConfigFile.J2.LicensePlate,
		ConfigFile.J2.Route,
		ConfigFile.J2.Reference,
		ConfigFile.J2.TargetCompany,
		ConfigFile.J2.Company,
		ConfigFile.J2.Postpaid,
		ConfigFile.J2.J2Postpaid,
		ConfigFile.J2.TotalFee,
	}
	var startIdx = ConfigFile.J2.StartIdx
	j2TitleIdx := getCellTitle(j2Sheet, j2Titles[:], startIdx)

	noIdx := j2TitleIdx[0]
	dateIdx := j2TitleIdx[1]
	licenceIdx := j2TitleIdx[2]
	routeIdx := j2TitleIdx[3]
	referenceIdx := j2TitleIdx[4]
	targetCompanyIdx := j2TitleIdx[5]
	companyIdx := j2TitleIdx[6]
	postPaidIdx := j2TitleIdx[7]
	j2PostPaidIdx := j2TitleIdx[8]
	totalFeeIdx := j2TitleIdx[9]

	result := make(map[models.SheetComp]models.CompReturn)
	for idx := 0 + startIdx + 1; idx < j2Sheet.MaxRow; idx++ {
		// 청구업체
		targetCompanyCell, _ := j2Sheet.Cell(idx, targetCompanyIdx)
		targetCompany := targetCompanyCell.String()

		if value, exists := ConfigFile.Company[companyFilter]; exists && len(value) > 0 && !Contains(value, targetCompany) {
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

		// 소속
		companyCell, _ := j2Sheet.Cell(idx, companyIdx)
		company := companyCell.String()

		// 후불
		postPaidCell, _ := j2Sheet.Cell(idx, postPaidIdx)
		postPaid := postPaidCell.String()
		postPaid = strings.Trim(postPaid, " ")

		// J2후불
		j2PostPaidCell, _ := j2Sheet.Cell(idx, j2PostPaidIdx)
		j2PostPaid := j2PostPaidCell.String()
		j2PostPaid = strings.Trim(j2PostPaid, " ")

		// 청구
		totalFeeCell, _ := j2Sheet.Cell(idx, totalFeeIdx)
		totalFeeStr := totalFeeCell.String()
		totalFeeStr = strings.Replace(totalFeeStr, ",", "", -1)
		totalFeeStr = strings.Trim(totalFeeStr, " ")
		tempTotalFee, err := strconv.Atoi(totalFeeStr)
		totalFee := -1
		if err == nil {
			totalFee = tempTotalFee
		}

		stage := 0
		if CheckStageSecond(company, postPaid, j2PostPaid) {
			stage = 2
		} else {
			stage = 1
		}

		if len(slice) < 2 {
			continue
		}

		sourceLayover := slice[0]
		dest := slice[1]
		isGansun := false
		isGansunOneway := false

		// 간선편도
		if strings.Contains(dest, "간선편도") {
			isGansun = true
			isGansunOneway = true
			dest = strings.Replace(dest, "간선편도대체", "", -1) // 간선편도대체 제거
			dest = strings.Replace(dest, "간선편도", "", -1)   // 간선편도 제거
		}

		// 고정간선
		if strings.Contains(dest, "간선") {
			isGansun = true
			isGansunOneway = false
			dest = strings.Replace(dest, "간선대체", "", -1) // 간선대체 제거
			dest = strings.Replace(dest, "간선", "", -1)   // 간선 제거
		}

		sliceLayover := strings.Split(sourceLayover, "/") // Split Layover
		source := sliceLayover[0]
		source = checkLayover(source)

		// 시프트면 하루 + 1, 토요일이면 + 2
		if strings.Contains(reference, "시프트") && !isGansun {
			if date.Weekday() == time.Saturday {
				date = date.AddDate(0, 0, 2)
			} else {
				date = date.AddDate(0, 0, 1)
			}
		}

		if strings.Contains(reference, ",") {
			reference = strings.Replace(reference, ",", "", -1)
			reference = "\"" + reference + "\""
		}

		if no == "" && date.String() == "" && licensePlate == "" && source == "" && dest == "" {
			break
		}

		// 간선, CJ 비교하는지 확인
		if (!GansunContain && isGansun) || (!CjContain && !isGansun) {
			continue
		}

		if checkGetData(source, dest, parseType) {
			each := new(models.SheetComp)
			each.Date = date.String()
			each.LicensePlate = licensePlate
			each.Source = source
			each.Destination = dest
			each.Gansun = isGansun
			each.GansunOneWay = isGansunOneway

			if _, exists := result[*each]; !exists {
				value := new(models.CompReturn)
				result[*each] = *value
			}

			value := result[*each]
			value.Idx = append(value.Idx, no)
			if reference != "" {
				value.Reference = append(value.Reference, reference)
			}
			value.Stage = stage
			if totalFee != -1 {
				value.TotalFee = append(value.TotalFee, totalFee)
			} else {
				value.TotalFee = append(value.TotalFee, 0)
			}

			// 간선일때
			if isGansun {
				// 특정 경로이고
				if each, exists := GansunFile.Route[source]; exists && Contains(each[0], dest) {
					// 청구업체가 틀리면 기록
					if targetCompany != each[1][0] {
						value.Reference = append(value.Reference, "간선 청구업체 오류: "+targetCompany+" 예상: "+each[1][0])
					}
				}
			}

			result[*each] = value
		}
	}

	return result
}

func ParseCJData(cjSheet *xlsx.Sheet, parseType string) map[models.SheetComp]models.CompReturn {
	var cjTitles = [...]string{
		ConfigFile.Cj.No, ConfigFile.Cj.Date,
		ConfigFile.Cj.LicensePlate,
		ConfigFile.Cj.Source,
		ConfigFile.Cj.Destination,
		ConfigFile.Cj.Reference,
		ConfigFile.Cj.DetourFeeType,
		ConfigFile.Cj.DetourFee,
		ConfigFile.Cj.DetourFeeType3,
		ConfigFile.Cj.DetourFee3,
		ConfigFile.Cj.MultiTourPercent,
		ConfigFile.Cj.TotalFee,
	}
	cjTitleIdx := getCellTitle(cjSheet, cjTitles[:], 0)
	var startIdx = ConfigFile.Cj.StartIdx

	noIdx := cjTitleIdx[0]
	dateIdx := cjTitleIdx[1]
	licenceIdx := cjTitleIdx[2]
	sourceIdx := cjTitleIdx[3]
	destIdx := cjTitleIdx[4]
	referenceIdx := cjTitleIdx[5]
	detourFeeTypeIdx := cjTitleIdx[6]
	detourFeeIdx := cjTitleIdx[7]
	detourFeeType3Idx := cjTitleIdx[8]
	detourFee3Idx := cjTitleIdx[9]
	multiTourPercentIdx := cjTitleIdx[10]
	totalFeeIdx := cjTitleIdx[11]

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
		source = strings.Replace(source, " ", "", -1)   // Trim
		source = strings.Replace(source, "Sub", "", -1) // Sub 제거
		source = strings.Replace(source, "Hub", "", -1) // Hub 제거
		source = strings.Replace(source, "콘솔", "", -1)  // 콘솔 제거
		source = checkLayover(source)

		// 도착
		destCell, _ := cjSheet.Cell(idx, destIdx)
		dest := destCell.String()
		dest = strings.Replace(dest, " ", "", -1)   // Trim
		dest = strings.Replace(dest, "Sub", "", -1) // Sub 제거
		dest = strings.Replace(dest, "Hub", "", -1) // Hub 제거
		dest = strings.Replace(dest, "콘솔", "", -1)  // 콘솔 제거
		dest = strings.Replace(dest, "지역", "", -1)  // 지역 제거

		// 비고
		referenceCell, _ := cjSheet.Cell(idx, referenceIdx)
		reference := referenceCell.String()
		if strings.Contains(reference, ",") {
			reference = strings.Replace(reference, ",", "", -1)
			reference = "\"" + reference + "\""
		}

		// 총운송비용
		totalFeeCell, _ := cjSheet.Cell(idx, totalFeeIdx)
		totalFeeStr := totalFeeCell.String()
		totalFeeStr = strings.Replace(totalFeeStr, ",", "", -1)
		totalFeeStr = strings.Trim(totalFeeStr, " ")
		tempTotalFee, err := strconv.Atoi(totalFeeStr)
		totalFee := -1
		if err == nil {
			totalFee = tempTotalFee
		}

		// 추가운임구분
		detourFeeTypeCell, _ := cjSheet.Cell(idx, detourFeeTypeIdx)
		detourFeeType := detourFeeTypeCell.String()

		// 추가운임
		detourFeeCell, _ := cjSheet.Cell(idx, detourFeeIdx)
		detourFee := 0
		if detourFeeCell.String() != "" || detourFeeCell.String() != "0" {
			temp, err := strconv.Atoi(detourFeeCell.String())
			if err == nil {
				detourFee = temp
			}
		}

		// 추가운임구분3
		detourFeeType3Cell, _ := cjSheet.Cell(idx, detourFeeType3Idx)
		detourFeeType3 := detourFeeType3Cell.String()

		// 추가운임3
		detourFee3Cell, _ := cjSheet.Cell(idx, detourFee3Idx)
		detourFee3 := 0
		if detourFee3Cell.String() != "" || detourFee3Cell.String() != "0" {
			temp, err := strconv.Atoi(detourFee3Cell.String())
			if err == nil {
				detourFee3 = temp
			}
		}

		// 다회전기준요율
		multiTourPercentCell, _ := cjSheet.Cell(idx, multiTourPercentIdx)
		multiTourPercent := multiTourPercentCell.String()

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
			each.GansunOneWay = false

			if _, exists := result[*each]; !exists {
				value := new(models.CompReturn)
				result[*each] = *value
			}

			value := result[*each]
			value.Idx = append(value.Idx, no)
			if reference != "" {
				value.Reference = append(value.Reference, reference)
			}
			if detourFeeType != "" {
				value.DetourFeeType = append(value.DetourFeeType, detourFeeType)
			}
			if detourFeeType3 != "" {
				value.DetourFeeType3 = append(value.DetourFeeType3, detourFeeType3)
			}
			if multiTourPercent != "" {
				value.MultiTourPercent = append(value.MultiTourPercent, multiTourPercent)
			}
			value.DetourFee += detourFee
			value.DetourFee3 += detourFee3
			if totalFee != -1 {
				value.TotalFee = append(value.TotalFee, totalFee)
			} else {
				value.TotalFee = append(value.TotalFee, 0)
			}

			result[*each] = value
		}
	}

	return result

}

func ParseGansunData(gansunSheet *xlsx.Sheet, parseType string) map[models.SheetComp]models.CompReturn {
	var gansunTitles = [...]string{
		ConfigFile.Gansun.No,
		ConfigFile.Gansun.Date,
		ConfigFile.Gansun.LicensePlate,
		ConfigFile.Gansun.Source,
		ConfigFile.Gansun.Destination,
		ConfigFile.Gansun.Reference,
		ConfigFile.Gansun.CarType,
		ConfigFile.Gansun.DetourFeeType,
		ConfigFile.Gansun.DetourFee,
		ConfigFile.Gansun.DetourFeeType3,
		ConfigFile.Gansun.DetourFee3,
		ConfigFile.Gansun.MultiTourPercent,
		ConfigFile.Gansun.TotalFee,
	}
	gansunTitleIdx := getCellTitle(gansunSheet, gansunTitles[:], 0)
	var startIdx = ConfigFile.Gansun.StartIdx

	noIdx := gansunTitleIdx[0]
	dateIdx := gansunTitleIdx[1]
	licenceIdx := gansunTitleIdx[2]
	sourceIdx := gansunTitleIdx[3]
	destIdx := gansunTitleIdx[4]
	referenceIdx := gansunTitleIdx[5]
	carTypeIdx := gansunTitleIdx[6]
	detourFeeTypeIdx := gansunTitleIdx[7]
	detourFeeIdx := gansunTitleIdx[8]
	detourFeeType3Idx := gansunTitleIdx[9]
	detourFee3Idx := gansunTitleIdx[10]
	multiTourPercentIdx := gansunTitleIdx[11]
	totalFeeIdx := gansunTitleIdx[12]

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
		source = strings.Replace(source, " ", "", -1) // Trim

		if source == "이천MP" {
			source = strings.Replace(source, "이천MP", "이천", -1) // 이천MP -> 이천
		}

		source = strings.Replace(source, "Sub", "", -1) // Sub 제거
		source = strings.Replace(source, "Hub", "", -1) // Hub 제거
		source = strings.Replace(source, "콘솔", "", -1)  // 콘솔 제거
		source = checkLayover(source)

		// 도착
		destCell, _ := gansunSheet.Cell(idx, destIdx)
		dest := destCell.String()
		dest = strings.Replace(dest, " ", "", -1)   // Trim
		dest = strings.Replace(dest, "Sub", "", -1) // Sub 제거
		dest = strings.Replace(dest, "Hub", "", -1) // Hub 제거
		dest = strings.Replace(dest, "콘솔", "", -1)  // 콘솔 제거

		// 차량 구분
		carTypeCell, _ := gansunSheet.Cell(idx, carTypeIdx)
		carType := carTypeCell.String()
		isGansunOneway := false
		if strings.Contains(carType, "편도") {
			isGansunOneway = true
		}

		// 비고
		referenceCell, _ := gansunSheet.Cell(idx, referenceIdx)
		reference := referenceCell.String()
		if strings.Contains(reference, ",") {
			reference = strings.Replace(reference, ",", "", -1)
			reference = "\"" + reference + "\""
		}

		// 총운송비용
		totalFeeCell, _ := gansunSheet.Cell(idx, totalFeeIdx)
		totalFeeStr := totalFeeCell.String()
		totalFeeStr = strings.Replace(totalFeeStr, ",", "", -1)
		totalFeeStr = strings.Trim(totalFeeStr, " ")
		tempTotalFee, err := strconv.Atoi(totalFeeStr)
		totalFee := -1
		if err == nil {
			totalFee = tempTotalFee
		}

		// 추가운임구분
		detourFeeTypeCell, _ := gansunSheet.Cell(idx, detourFeeTypeIdx)
		detourFeeType := detourFeeTypeCell.String()

		// 추가운임
		detourFeeCell, _ := gansunSheet.Cell(idx, detourFeeIdx)
		detourFee := 0
		if detourFeeCell.String() != "" || detourFeeCell.String() != "0" {
			temp, err := strconv.Atoi(detourFeeCell.String())
			if err == nil {
				detourFee = temp
			}
		}

		// 추가운임구분3
		detourFeeType3Cell, _ := gansunSheet.Cell(idx, detourFeeType3Idx)
		detourFeeType3 := detourFeeType3Cell.String()

		// 추가운임3
		detourFee3Cell, _ := gansunSheet.Cell(idx, detourFee3Idx)
		detourFee3 := 0
		if detourFee3Cell.String() != "" || detourFee3Cell.String() != "0" {
			temp, err := strconv.Atoi(detourFee3Cell.String())
			if err == nil {
				detourFee3 = temp
			}
		}

		// 다회전기준요율
		multiTourPercentCell, _ := gansunSheet.Cell(idx, multiTourPercentIdx)
		multiTourPercent := multiTourPercentCell.String()

		if no == "합계" || no == "" && date.String() == "" && licensePlate == "" && source == "" && dest == "" {
			break
		}

		if checkGetData(source, dest, parseType) {
			each := new(models.SheetComp)
			each.Date = date.String()
			each.LicensePlate = licensePlate
			each.Source = source
			each.Destination = dest
			each.Gansun = true
			each.GansunOneWay = isGansunOneway

			if _, exists := result[*each]; !exists {
				value := new(models.CompReturn)
				result[*each] = *value
			}

			value := result[*each]
			value.Idx = append(value.Idx, no)
			if reference != "" {
				value.Reference = append(value.Reference, reference)
			}
			if detourFeeType != "" {
				value.DetourFeeType = append(value.DetourFeeType, detourFeeType)
			}
			if detourFeeType3 != "" {
				value.DetourFeeType3 = append(value.DetourFeeType3, detourFeeType3)
			}
			if multiTourPercent != "" {
				value.MultiTourPercent = append(value.MultiTourPercent, multiTourPercent)
			}
			value.DetourFee += detourFee
			value.DetourFee3 += detourFee3
			if totalFee != -1 {
				value.TotalFee = append(value.TotalFee, totalFee)
			} else {
				value.TotalFee = append(value.TotalFee, 0)
			}

			result[*each] = value
		}
	}

	return result

}
