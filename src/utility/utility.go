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
