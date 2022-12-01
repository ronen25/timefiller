package main

import (
	"fmt"
	"math/rand"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Locations
const (
	LocationYear  = "B2"
	LocationMonth = "B3"

	LocationEmployeeName = "C5"

	StartLocationLine         = 8
	EndLocationLine           = 37
	LocationDay               = "B%d"
	StartLocationClockInHour  = "C%d"
	StartLocationClockOutHour = "D%d"
	LocationExcuse            = "G%d"
)

var BasePrefix = [...]string{"בעיקר", "דברים שונים אבל בעיקר", "דברים שונים +", "שונות +"}
var BaseExcuses = [...]string{"פיתוח", "בדיקות", "בדיקת באג דחוף", "אפיון", "אפיון פיצ׳ר", "פיתוח פיצ׳ר"}

func GenerateExcuse(cellIndex int) string {
	excuseIndex := rand.Intn(len(BaseExcuses))

	if cellIndex%10 == 0 {
		prefixIndex := rand.Intn(len(BasePrefix))

		return fmt.Sprintf("%s %s", BasePrefix[prefixIndex], BaseExcuses[excuseIndex])
	}

	return BaseExcuses[excuseIndex]
}

func IsApplicableDay(dayValue string) (bool, error) {
	date, parseErr := time.Parse("01-02-06", dayValue)
	if parseErr != nil {
		return false, parseErr
	}

	return date.Weekday() != time.Friday && date.Weekday() != time.Saturday, nil
}

func FillFile(path string, config *Config) (*excelize.File, error) {
	rawFile, err := os.OpenFile(path, os.O_RDONLY, 0644)
	if err != nil {
		return nil, err
	}
	defer rawFile.Close()

	file, err := excelize.OpenReader(rawFile)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	currentDate := time.Now()
	month := int(currentDate.Month()) - 1
	if month == 0 {
		// Got January so roll it back to 1
		month = 1
	}

	sheetName := file.GetSheetName(file.GetActiveSheetIndex())
	file.SetCellValue(sheetName, LocationYear, currentDate.Year())
	file.SetCellValue(sheetName, LocationMonth, month)

	file.SetCellValue(sheetName, LocationEmployeeName, config.EmployeeName)

	for i := StartLocationLine; i <= EndLocationLine; i++ {
		// Check if the day is a friday or a saturday, if it is, skip it.
		cellDate := fmt.Sprintf(LocationDay, i)
		dateRaw, dateErr := file.GetCellValue(sheetName, cellDate)
		if dateErr != nil {
			return nil, dateErr
		}

		isApplicable, parseErr := IsApplicableDay(dateRaw)
		if parseErr != nil {
			return nil, parseErr
		}

		if !isApplicable {
			continue
		}

		clockInHour := fmt.Sprintf(StartLocationClockInHour, i)
		clockOutHour := fmt.Sprintf(StartLocationClockOutHour, i)
		excuse := fmt.Sprintf(LocationExcuse, i)

		file.SetCellValue(sheetName, clockInHour, "9:00")
		file.SetCellValue(sheetName, clockOutHour, "18:30")
		file.SetCellValue(sheetName, excuse, GenerateExcuse(i))
	}

	fileExt := filepath.Ext(path)
	fileBase := strings.TrimSuffix(path, fileExt)
	copyPath := fmt.Sprintf("%s_%s_%d_%d.xlsx", fileBase, config.EmployeeName,
		month, currentDate.Year())
	err = file.SaveAs(copyPath)
	if err != nil {
		return nil, err
	}

	return file, nil
}
