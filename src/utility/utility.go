package utility

import (
	"bufio"
	"fmt"
	"os"
	"sort"
	"strings"

	"js.comp.dispatching/src/models"
)

func GetParameters() (j2SheetName, cjSheetName, gansunSheetName, resultFileName, resultType, companyFilter, errCode string) {
	fmt.Print("J2 파일 시트명(Default: sheet1): ")
	j2Buf := bufio.NewScanner(os.Stdin)
	j2Buf.Scan()
	j2SheetName = j2Buf.Text()

	if CjContain {
		fmt.Print("CJ 파일 시트명(Default: sheet1): ")
		cjBuf := bufio.NewScanner(os.Stdin)
		cjBuf.Scan()
		cjSheetName = cjBuf.Text()
	}

	if GansunContain {
		fmt.Print("간선 파일 시트명(Default: sheet1): ")
		gansunBuf := bufio.NewScanner(os.Stdin)
		gansunBuf.Scan()
		gansunSheetName = gansunBuf.Text()
	}

	fmt.Print("결과 파일명(Default: result.csv): ")
	resBuf := bufio.NewScanner(os.Stdin)
	resBuf.Scan()
	resultFileName = resBuf.Text()

	fmt.Print("노선 필터(TOML파일 target): ")
	fmt.Scanln(&resultType)

	fmt.Print("업체 필터(TOML파일 company): ")
	fmt.Scanln(&companyFilter)

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

	if _, exists := ConfigFile.Target[resultType]; !exists {
		resultType = "1"
	}

	if _, exists := ConfigFile.Company[companyFilter]; !exists {
		errCode = models.InvalidCompany
	}

	return
}

// Contains is check the substr in array s.
func Contains(s []string, substr string) bool {
	for _, v := range s {
		if v == substr {
			return true
		}
	}
	return false
}

// GetArrayData is make string about array data arr.
func GetArrayData(arr []string, split string) string {
	if len(arr) == 0 {
		return ""
	}

	no := arr[0]

	for idx := 1; idx < len(arr); idx++ {
		no = no + split + arr[idx]
	}
	return no
}

// GetStage returns stage string data.
func GetStage(stage int) string {
	switch stage {
	case 0:
		return "0단계완료"
	case 1:
		return "1단계완료"
	case 2:
		return "2단계완료"
	case 3:
		return "3단계완료"
	}

	return ""
}

// CheckStageSecond is function that check the stage is two.
func CheckStageSecond(company, postPaid, j2PostPaid string) bool {
	return (company == "제이투" && postPaid != "" && j2PostPaid != "") || (company != "제이투" && postPaid != "")
}

// CheckTotalFee is function that check the total fee.
func CheckTotalFee(first, second []int) (isSame bool, firstTotal, secondTotal int) {
	// firstTotal = ArraySum(first)
	// secondTotal = ArraySum(second)

	if len(first) != len(second) {
		return
	}

	sort.Slice(first, func(i, j int) bool {
		return i < j
	})

	sort.Slice(second, func(i, j int) bool {
		return i < j
	})

	isSame = true
	for idx := 0; idx < len(first); idx++ {
		if first[idx] != second[idx] {
			isSame = false
			return
		}
	}

	return
}

func ArraySum(arr []int) (sum int) {
	for _, value := range arr {
		sum += value
	}

	return
}

// 가져오는 데이터가 맞는지 조건 확인
func checkGetData(source, dest, parseType string) bool {
	for _, target := range ConfigFile.Target[parseType] {
		if strings.Contains(source, target) || strings.Contains(dest, target) {
			return true
		}
	}

	return false
}

// Layover Check
func checkLayover(key string) string {
	if value, exists := ConfigFile.Direct[key]; exists {
		return value
	}
	return key
}

func CheckStage(stage, j2, cj int) int {
	if stage == 2 && j2 == cj {
		return 3
	}
	return stage
}
