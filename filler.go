package main

import (
	"fmt"
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
	StartLocationClockInHour  = "C%d"
	StartLocationClockOutHour = "D%d"
	LocationExcuse            = "G%d"
)

func CopyFileToPath(originalPath, suffix string) (string, error) {
	input, err := os.ReadFile(originalPath)
	if err != nil {
		return "", err
	}

	fileExt := filepath.Ext(originalPath)
	fileBase := strings.TrimSuffix(originalPath, fileExt)
	copyPath := fmt.Sprintf("%s_%s.%s", fileBase, suffix, originalPath)

	err = os.WriteFile(copyPath, input, 0644)
	if err != nil {
		return "", nil
	}

	return copyPath, nil
}

func FillFile(path string, config *Config) (*excelize.File, error) {
	rawFile, err := os.OpenFile(path, os.O_RDONLY, 0644)
	if err != nil {
		return nil, err
	}

	file, err := excelize.OpenReader(rawFile)
	if err != nil {
		return nil, err
	}

	currentDate := time.Now()

	sheetName := file.GetSheetName(file.GetActiveSheetIndex())
	file.SetCellValue(sheetName, LocationYear, currentDate.Year())
	file.SetCellValue(sheetName, LocationMonth, currentDate.Month())

	file.SetCellValue(sheetName, LocationEmployeeName, config.EmployeeName)

	for i := StartLocationLine; i <= EndLocationLine; i++ {
		clockInHour := fmt.Sprintf(StartLocationClockInHour, i)
		clockOutHour := fmt.Sprintf(StartLocationClockOutHour, i)
		excuse := fmt.Sprintf(LocationExcuse, i)

		file.SetCellValue(sheetName, clockInHour, "9:00")
		file.SetCellValue(sheetName, clockOutHour, "18:30")
		file.SetCellValue(sheetName, excuse, "עבודת פיתוח או משהו לא יודע")
	}

	fileExt := filepath.Ext(path)
	fileBase := strings.TrimSuffix(path, fileExt)
	copyPath := fmt.Sprintf("%s_%s.%s", fileBase, "filled", path)
	err = file.SaveAs(copyPath)
	if err != nil {
		return nil, err
	}

	return file, nil
}
