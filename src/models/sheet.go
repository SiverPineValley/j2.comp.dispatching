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
	Idx              []string
	SourcePostfix    string
	Reference        []string
	DetourFeeType    []string
	DetourFee        []string
	DetourFeeType3   []string
	DetourFee3       []string
	DetourFair       []string
	MultiTourPercent []string
	Stage            int
	TotalFee         []int
}

type CompData struct {
	Date             string
	LicensePlate     string
	Source           string
	Destination      string
	IsGansun         bool
	IsGansunOneway   bool
	J2Reference      string
	CJReference      string
	J2No             []string
	CJNo             []string
	GansunNo         []string
	DetourFeeType    string
	DetourFee        string
	DetourFeeType3   string
	DetourFee3       string
	DetourFair       string
	MultiTourPercent string
	J2               bool
	CJ               bool
	Gansun           bool
	Stage            int
	FirstTotalFee    int
	SecondTotalFee   int
}
