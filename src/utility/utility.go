package utility

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
