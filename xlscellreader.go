package xlscellreader

import (
	"errors"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

const CELL_TYPE_UNSET = 0

func NewCellReader(f *excelize.File) CellReader {
	return CellReader{f}
}

type CellReader struct {
	f *excelize.File
}

/*
Reads a string at the sheet and axis location
*/
func (c CellReader) GetString(sheet string, axis string) (string, error) {
	return c.f.GetCellValue(sheet, axis)
}

/*
Reads an int at the sheet and axis location.
Returns an error if the cell is not numeric or the string value cannot be parsed to an integer.
*/
func (c CellReader) GetInt(sheet string, axis string) (int, error) {
	val, err := c.getVal(sheet, axis, excelize.CellTypeNumber)
	if err != nil {
		return 0, err
	}
	return strconv.Atoi(val)
}

/*
Reads an int at the sheet and axis location.
Returns an error if the cell is not numeric or the string value cannot be parsed to a float64.
*/
func (c CellReader) GetFloat(sheet string, axis string) (float64, error) {
	val, err := c.getVal(sheet, axis, excelize.CellTypeNumber)
	if err != nil {
		return 0, err
	}
	return strconv.ParseFloat(val, 64)
}

/*
Reads a Date/Time at the sheet and axis location.
Returns an error if the cell is not numeric or the string value
cannot be parsed to an int and converted into a date value.
*/
func (c CellReader) GetDate(sheet string, axis string) (time.Time, error) {
	var xlsEpoch = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	c.f.SetCellStyle(sheet, axis, axis, 0)
	val, err := c.getVal(sheet, axis, excelize.CellTypeNumber)
	if err != nil {
		return time.Time{}, err
	}
	xlsTime, err := strconv.Atoi(val)
	if err != nil {
		return time.Time{}, err
	}
	return xlsEpoch.Add(time.Second * time.Duration(xlsTime*86400)), nil
}

/*
Reads a Date/Time at the sheet and axis location by parsing the
formatted date representation in excel.
Returns an error if the cell is not numeric or the string value
cannot be parsed to a Date/Time.
*/
func (c CellReader) GetFormattedDate(sheet string, axis string, dateFmt string) (time.Time, error) {
	val, err := c.getVal(sheet, axis, excelize.CellTypeNumber)
	if err != nil {
		return time.Time{}, err
	}
	return time.Parse(dateFmt, val)
}

func (c CellReader) getVal(sheet string, axis string, cellType excelize.CellType) (string, error) {
	cellTypeRead, err := c.f.GetCellType(sheet, axis)
	if err != nil {
		return "", err
	}
	if cellTypeRead != CELL_TYPE_UNSET && cellTypeRead != cellType {
		return "", errors.New("invalid cell type.")
	}
	return c.f.GetCellValue(sheet, axis)
}
