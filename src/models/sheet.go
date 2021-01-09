package models

const (
	ParseTypeIcheon string = "1"
	ParseTypeAll    string = "2"
)

type SheetComp struct {
	Date         string
	LicensePlate string
	Source       string
	Destination  string
}

type CompReturn struct {
	Idx       []string
	Reference []string
}

type CompData struct {
	Date         string
	LicensePlate string
	Source       string
	Destination  string
	Reference    []string
	J2No         []string
	CJNo         []string
	J2           bool
	CJ           bool
}
