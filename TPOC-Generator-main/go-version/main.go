package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type HeaderMap struct {
	cord_x          int
	cord_y          int
	fbs_area        int
	fbs_process     int
	fbs_unit        int
	fbs_unit_number int
}

type TagInfo struct {
	cord_x          string
	cord_y          string
	fbs_area        string
	fbs_process     string
	fbs_unit        string
	fbs_unit_number string
}

var xCreationPoint int = -1000
var yCreationPoint int = -1000

func main() {

	f, err := excelize.OpenFile("FBS_Data.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := f.GetRows("Properties")

	hMap := HeaderMap{}

	for i, header := range rows[0] {
		if header == "OffsetX" {
			hMap.cord_x = i
		} else if header == "OffsetY" {
			hMap.cord_y = i
		} else if header == "Custom.FBS_Area" {
			hMap.fbs_area = i
		} else if header == "Custom.FBS_Process" {
			hMap.fbs_process = i
		} else if header == "Custom.FBS_Unit" {
			hMap.fbs_unit = i
		} else if header == "Custom.Unit Number" {
			hMap.fbs_unit_number = i
		}
	}

	dataValue := []TagInfo{}

	for _, row := range rows[1:] {

		//Excel from autocad divides x,y,z cord by 10

		x, _ := strconv.Atoi(row[hMap.cord_x])
		x = x * 10

		y, _ := strconv.Atoi(row[hMap.cord_x])
		y = y * 10

		d := TagInfo{
			cord_x:          strconv.Itoa(x),
			cord_y:          strconv.Itoa(y),
			fbs_area:        row[hMap.fbs_area],
			fbs_process:     row[hMap.fbs_process],
			fbs_unit:        row[hMap.fbs_unit],
			fbs_unit_number: row[hMap.fbs_unit_number],
		}

		dataValue = append(dataValue, d)

	}

	f_out, err := os.Create("./FBS_Tags.txt")

	w := bufio.NewWriter(f_out)

	_, err = fmt.Fprintf(w, "(command) ATTDIA 0 ATTREQ 0\n")
	_, err = fmt.Fprintf(w, "(command) -DWGUNITS 3 2 2 Y Y\n")

	for _, d := range dataValue {
		_, err = fmt.Fprintf(w, "(command) -TEXT -1000,-1000,0 250 0 %s.%s.%s%s\n",
			d.fbs_area,
			d.fbs_process,
			d.fbs_unit,
			d.fbs_unit_number)

		_, err = fmt.Fprintf(w, "(command) MOVE -999,-999,0  %d,%d,0 %s,%s,0\n",
			xCreationPoint,
			yCreationPoint,
			d.cord_x,
			d.cord_y)

		//(command) MOVE -999,-999,0  -1000,-1000,0 2000,2000,0
	}

	_, err = fmt.Fprintf(w, "(command) ATTDIA 1 ATTREQ 1\n")

	w.Flush()
	f_out.Close()

	err = os.Rename("./FBS_Tags.txt", "./FBS_Tags.scr")
	if err != nil {
		fmt.Println(err)
	}

}
