package main

import (
	"github.com/gin-gonic/gin"
	"github.com/tealeg/xlsx"
	"fmt"
	"net/http"
	"log"
	"strconv"
	"os"
	"io"
	"time"
	// "reflect"
	// "io/ioutil"
)

// TimeSheet for sheet slsx
type TimeSheet struct {
	Date string `json:"date"`
	Name string `json:"name"`
	Code string `json:"code"`
	Rule string `json:"rule"`
	TimeIn string `json:"time_in"`
	TimeOut string `json:"time_out"`
} 

// CORSMiddleware enable cors
func CORSMiddleware() gin.HandlerFunc {
	return func(c *gin.Context) {
		c.Writer.Header().Set("Access-Control-Allow-Origin", "*")
		c.Writer.Header().Set("Access-Control-Allow-Credentials", "true")
		c.Writer.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		c.Writer.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS, GET, PUT")
		if c.Request.Method == "OPTIONS" {
				c.AbortWithStatus(204)
				return
		}
		c.Next()
	}
}

func isExistInArray(item TimeSheet, list []TimeSheet) (bool){
	var result bool
	for _, value := range list {
		if (value.Name == item.Name && value.Date == item.Date ) {
			result = true
			break
		}
	}
	return result
}


func findIndexInArray(item TimeSheet, list []TimeSheet) (int){
	var result int
	for index, value := range list {
		if (value.Name == item.Name && value.Date == item.Date ) {
			result = index
			break
		}
	}
	return result
}

func countInArray(item TimeSheet, list []TimeSheet) int {
	var result = 0
	for _, value := range list {
		if (value.Name == item.Name && value.Date == item.Date ) {
			result = result+1
		}
	}
	return result
}

func main() {
	router := gin.Default()
	// Set a lower memory limit for multipart forms (default is 32 MiB)
	// router.MaxMultipartMemory = 8 << 20  // 8 MiB
	router.Use(CORSMiddleware())
	router.POST("/upload", func(c *gin.Context) {
		// single file
		fileImport, _ := c.FormFile("file")
		log.Println(fileImport.Filename)
		fileOpen, err := fileImport.Open()

		if err != nil {
			c.String(http.StatusOK, fmt.Sprintf("'%s' loi  o day ne!",err))
    }
		xlFile, err := xlsx.OpenReaderAt(fileOpen, fileImport.Size)
    if err != nil {
			c.String(http.StatusOK, fmt.Sprintf("'%s' loi cmnr!",err))
		}
		var DataTimeSheet = []TimeSheet{}
		
		// read file slxs 
    for _, sheet := range xlFile.Sheets {
			for indexRow, row := range sheet.Rows {
					if indexRow >=10 {
						var TimeSheetRow = TimeSheet{}
						value, _ := row.Cells[1].FormattedValue()
						TimeSheetRow.Date = value
						TimeSheetRow.Code = row.Cells[2].Value
						TimeSheetRow.Name = row.Cells[3].Value
						TimeSheetRow.Rule = row.Cells[4].Value
						TimeSheetRow.TimeIn = row.Cells[5].Value
						TimeSheetRow.TimeOut = row.Cells[6].Value
						DataTimeSheet = append(DataTimeSheet, TimeSheetRow)
					}
			}
		}


		// lay cap gia tri dau tien va cuoi cung
		var DataRender = []TimeSheet{}
		for _, item := range DataTimeSheet {
			value := isExistInArray(item, DataRender)
			count := countInArray(item, DataRender)


			if count <= 1 {
				DataRender = append(DataRender, item)
				continue
			}
			
			if value && count > 1 {
				DataRender = DataRender[:len(DataRender)-1]
				DataRender = append(DataRender, item)
			}
		}


		// lay gio dau tien va gio cuoi cung
		var DataExcel = []TimeSheet{}
		for _, item := range DataRender {
			count := countInArray(item, DataExcel)

			if count == 0 {
				var TimeSheetRow = TimeSheet{}
				TimeSheetRow.Date = item.Date
				TimeSheetRow.Code = item.Code
				TimeSheetRow.Name = item.Name
				TimeSheetRow.Rule = item.Rule
				if item.TimeIn == "" {
					TimeSheetRow.TimeIn = item.TimeOut
				} else {
					TimeSheetRow.TimeIn = item.TimeIn
				}
				TimeSheetRow.TimeOut = item.TimeOut
				DataExcel = append(DataExcel, TimeSheetRow)
				continue
			}
			
			if count == 1 {
				indexFound :=findIndexInArray(item, DataExcel)
				if item.TimeOut == "" {
					DataExcel[indexFound].TimeOut = item.TimeIn
				} else {
					DataExcel[indexFound].TimeOut = item.TimeOut
				}
			}
		}


		// export excel
		var file *xlsx.File
    var sheet *xlsx.Sheet
    var row *xlsx.Row

    file = xlsx.NewFile()
    sheet, err = file.AddSheet("Sheet1")
    if err != nil {
        fmt.Printf(err.Error())
		}
		for index, item := range DataExcel { 
			row = sheet.AddRow()
			indexCell := row.AddCell()
			dateCell := row.AddCell()
			codeCell := row.AddCell()
			nameCell := row.AddCell()
			ruleCell := row.AddCell()
			timeInCell := row.AddCell()
			timeOutCell := row.AddCell()
			indexCell.Value = strconv.Itoa(index)
			dateCell.Value = item.Date
			nameCell.Value = item.Name
			codeCell.Value = item.Code
			ruleCell.Value = item.Rule
			timeInCell.Value = item.TimeIn
			timeOutCell.Value = item.TimeOut
		}
		// render name
		t := time.Now()
		formatedTime := t.Format(time.RFC3339)
		fileName := formatedTime+".xlsx"
		err = file.Save(fileName)
		
    if err != nil {
        fmt.Printf(err.Error())
		}

		var r io.Reader
		r, err = os.Open(fileName)
		fileExport, err := os.Open(fileName)
		if err != nil {
				log.Fatal(err)
		}
	
		fi, err := fileExport.Stat()
		contentLength := int64(fi.Size())
		contentType := "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		extraHeaders := map[string]string{
			"Content-Disposition": `attachment; filename="`+fileName+`"`,
		}

		c.DataFromReader(http.StatusOK, contentLength, contentType, r, extraHeaders)
	})


	router.Run(":1234")
}

