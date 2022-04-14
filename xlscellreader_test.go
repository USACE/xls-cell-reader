package xlscellreader

import (
	"fmt"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

func TestCellReader(t *testing.T) {
	f, err := excelize.OpenFile("data/test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	cr := CellReader{f}

	vtest := 1
	v, err := cr.GetInt("Sheet1", "B1")
	if v != vtest {
		t.Fatalf("Expected %d got %d", vtest, v)
	}

	v2test := 2.123456
	v2, _ := cr.GetFloat("Sheet1", "B2")
	if v2 != v2test {
		t.Fatalf("Expected %f got %f", v2test, v2)
	}

	v3test, _ := time.Parse("2006-01-02", "2021-12-05")
	v3, _ := cr.GetFormattedDate("Sheet1", "B3", "01/02/06")
	if v3 != v3test {
		t.Fatalf("Expected %v got %v", v3test, v3)
	}

	v4, _ := cr.GetDate("Sheet1", "B3")
	if v4 != v3test {
		t.Fatalf("Expected %v got %v", v3test, v4)
	}

	v5test := "4"
	v5, _ := cr.GetString("Sheet1", "B4")
	if v5 != v5test {
		t.Fatalf("Expected %s got %s", v5test, v5)
	}

	v6test := false
	v6, _ := cr.GetBool("Sheet1", "B5")
	if v6 != v6test {
		t.Fatalf("Expected %t got %t", v6test, v6)
	}
}

func ExampleCellReader() {
	f, err := excelize.OpenFile("data/test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	cr := CellReader{f}

	v, err := cr.GetInt("Sheet1", "B1")
	if err != nil {
		fmt.Println(v)
	}

	v2, err := cr.GetFloat("Sheet1", "B2")
	if err != nil {
		fmt.Println(v2)
	}

	v3, err := cr.GetFormattedDate("Sheet1", "B3", "01/02/06")
	if err != nil {
		fmt.Println(v3)
	}

	v4, err := cr.GetDate("Sheet1", "B3")
	if err != nil {
		fmt.Println(v4)
	}

	v5, err := cr.GetString("Sheet1", "B4")
	if err != nil {
		fmt.Println(v5)
	}

}
