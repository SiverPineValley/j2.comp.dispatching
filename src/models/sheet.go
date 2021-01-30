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
	Gansun       bool
	GansunOneWay bool
}

type CompReturn struct {
	Idx          []string
	Reference    []string
	TrailerPlate []string
}

type CompData struct {
	Date           string
	LicensePlate   string
	Source         string
	Destination    string
	IsGansun       bool
	IsGansunOneway bool
	Reference      []string
	TrailerPlate   []string
	J2No           []string
	CJNo           []string
	GansunNo       []string
	J2             bool
	CJ             bool
	Gansun         bool
}
