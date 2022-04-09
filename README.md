## XLSCELLREADER
A utility to simplify reading of data from spreadsheets using excelize.  <br>
The excelize GetCellValue only returns string values from excel spreadsheets.  this utility wraps the excelize file and will return specific types. 

add to your go project with
```
go get github.com/usace/xlscellreader
```


```golang

import (
    "fmt"
    . "xlscellreader"
)

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
```