package utility

import (
	"bufio"
	"os"
	"strconv"

	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/transform"
	"js.comp.dispatching/src/models"
)

var CjContain bool = false
var GansunContain bool = false

func WriteXlsxData(comData []models.CompData) {
	// file := xlsx.NewFile()
}

func WriteCSVData(resultFileName string, comData []models.CompData) {
	// Comp 파일 생성
	resFile, err := os.Create("./" + resultFileName)
	if err != nil {
		panic(err)
	}

	// Comp csv writer 생성
	w := bufio.NewWriter(resFile)
	resWr := transform.NewWriter(w, korean.EUCKR.NewEncoder())

	// Comp 내용 쓰기
	if CjContain && GansunContain {
		resWr.Write([]byte("날짜, 차량번호, 출발, 도착, 간선여부, J2, CJ, Gansun, 톤수, 추가운임, 추가운임구분3, 추가운임3, 다회전기준요율, 경유비, 청구, " + ConfigFile.Cj.TotalFee + ", SJ비고, CJ비고, 완료단계\n"))
	} else if CjContain && !GansunContain {
		resWr.Write([]byte("날짜, 차량번호, 출발, 도착, 간선여부, J2, CJ, 톤수, 추가운임, 추가운임구분3, 추가운임3, 다회전기준요율, 경유비, 청구, " + ConfigFile.Cj.TotalFee + ", SJ비고, CJ비고, 완료단계\n"))
	} else if !CjContain && GansunContain {
		resWr.Write([]byte("날짜, 차량번호, 출발, 도착, 간선여부, J2, Gansun, 톤수, 추가운임, 추가운임구분3, 추가운임3, 다회전기준요율, 경유비, 청구, " + ConfigFile.Gansun.TotalFee + ", SJ비고, CJ비고, 완료단계\n"))
	}

	for _, value := range comData {
		gansun := ""
		if value.IsGansun && value.IsGansunOneway {
			gansun = "간선편도"
		} else if value.IsGansun && !value.IsGansunOneway {
			gansun = "고정간선"
		}

		// 자체 No, 날짜, 차량번호, 출발, 도착, 간선
		each := value.Date[0:10] + "," + value.LicensePlate + "," + value.Source + "," + value.Destination + "," + gansun

		// No 숫자, 비고
		if CjContain && GansunContain {
			if value.J2 && value.CJ {
				each = each + "," + strconv.Itoa(len(value.J2No)) + "," + strconv.Itoa(len(value.CJNo)) + ", 0"
			} else if value.J2 && value.Gansun {
				each = each + "," + strconv.Itoa(len(value.J2No)) + ", 0, " + strconv.Itoa(len(value.GansunNo))
			} else if value.J2 && !value.CJ && !value.Gansun {
				each = each + "," + strconv.Itoa(len(value.J2No)) + ", 0, 0"
			} else if !value.J2 && value.CJ && !value.Gansun {
				each = each + ",0," + strconv.Itoa(len(value.CJNo)) + ", 0"
			} else if !value.J2 && !value.CJ && value.Gansun {
				each = each + ",0, 0," + strconv.Itoa(len(value.GansunNo))
			}
		} else if CjContain && !GansunContain {
			if value.J2 && !value.CJ {
				each = each + "," + strconv.Itoa(len(value.J2No)) + ", 0"
			} else if !value.J2 && value.CJ {
				each = each + ",0," + strconv.Itoa(len(value.CJNo))
			} else {
				each = each + "," + strconv.Itoa(len(value.J2No)) + "," + strconv.Itoa(len(value.CJNo))
			}
		} else if !CjContain && GansunContain {
			if value.J2 && !value.Gansun {
				each = each + "," + strconv.Itoa(len(value.J2No)) + ", 0"
			} else if !value.J2 && value.Gansun {
				each = each + ",0," + strconv.Itoa(len(value.GansunNo))
			} else {
				each = each + "," + strconv.Itoa(len(value.J2No)) + "," + strconv.Itoa(len(value.GansunNo))
			}
		}

		each = each + "," + value.DetourFeeType + "," + value.DetourFee + "," + value.DetourFeeType3 + "," + value.DetourFee3 + "," + value.MultiTourPercent + "," + value.DetourFair + "," + strconv.Itoa(value.FirstTotalFee) + "," + strconv.Itoa(value.SecondTotalFee) + "," + value.J2Reference + "," + value.CJReference + "," + GetStage(value.Stage) + "\n"
		resWr.Write([]byte(each))
		w.Flush()
	}

	resWr.Close()
}
