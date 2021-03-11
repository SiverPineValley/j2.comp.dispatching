package utility

import "sort"

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
