package models

type SheetComp struct {
	Date         string
	LicensePlate string
	Source       string
	Destination  string
}

type CompData struct {
	Date         string
	LicensePlate string
	Source       string
	Destination  string
	J2No         []string
	CJNo         []string
	J2           bool
	CJ           bool
}
