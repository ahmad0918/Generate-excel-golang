package main

import (
	"github.com/xuri/excelize/v2"
	"fmt"
)

type GenerateExcelData struct {
	NamaPartner string
	UniqueCode  string
	KodeRef     string
	Link        string
}

func main() {
	dataArray := []GenerateExcelData{
		{
			NamaPartner: "XXXXXXX",
			UniqueCode:  "YYYYYYYYY",
			KodeRef:     "PDXXYYYY",
			Link:        "https://Example.link/PDXXYYYY",
		},
		{
			NamaPartner: "TokoPandai",
			UniqueCode:  "333972",
			KodeRef:     "PDTo3972",
			Link:        "https://Example.link/PDTo3972",
		},
		{
			NamaPartner: "TokoPandai",
			UniqueCode:  "234972",
			KodeRef:     "PDTo4972",
			Link:        "https://Example.link/PDTo4972",
		},
		{
			NamaPartner: "TokoPandai",
			UniqueCode:  "123472",
			KodeRef:     "PDTo3472",
			Link:        "https://Example.link/PDTo3472",
		},
		{
			NamaPartner: "GoTokopedia",
			UniqueCode:  "82110783987",
			KodeRef:     "PDGo3987",
			Link:        "https://Example.link/PDGo3987",
		},
		{
			NamaPartner: "Shopee",
			UniqueCode:  "D234h397",
			KodeRef:     "PDShh397",
			Link:        "https://Example.link/PDShh397",
		},
	}

	if err := GenerateExcel("Data-Excel.xlsx", dataArray); err != nil {
		fmt.Println(err, "error in Generate Report")
		return
	}

	fmt.Println("Success Generate Excel")
	return
}

func GenerateExcel(fileName string, models []GenerateExcelData) error {
	f := excelize.NewFile()
	defer closeExcelFile(f)

	sheetName := "Sheet1"
	index, err := createSheetWithHeaders(f, sheetName)
	if err != nil {
		return err
	}

	for i := 0; i < len(models); i++ {
		appendRow(f, models[i], i+2)
	}

	f.SetActiveSheet(index)

	filePath := "file_excel/" + fileName
	if err := saveExcelFile(f, filePath); err != nil {
		return err
	}

	return nil
}

func closeExcelFile(f *excelize.File) {
	if errFile := f.Close(); errFile != nil {
		fmt.Println(errFile, "[GenerateReportRefCode] Error in New File Excel")
	}
}

func createSheetWithHeaders(f *excelize.File, sheetName string) (int, error) {
	index, errSheet := f.NewSheet(sheetName)
	if errSheet != nil {
		fmt.Println(errSheet, "[GenerateReportRefCode] Error in New Sheet Excel")
		return 0, errSheet
	}

	col := map[string]string{
		"A1": "Nama Partner",
		"B1": "Unique Code",
		"C1": "Kode Referral",
		"D1": "Link Kode Referral",
	}

	for cell, value := range col {
		f.SetCellValue(sheetName, cell, value)
	}

	return index, nil
}

func appendRow(f *excelize.File, model GenerateExcelData, rowCount int) {
	f.SetCellValue("Sheet1", fmt.Sprintf("A%v", rowCount), model.NamaPartner)
	f.SetCellValue("Sheet1", fmt.Sprintf("B%v", rowCount), model.UniqueCode)
	f.SetCellValue("Sheet1", fmt.Sprintf("C%v", rowCount), model.KodeRef)
	f.SetCellValue("Sheet1", fmt.Sprintf("D%v", rowCount), model.Link)
}

func saveExcelFile(f *excelize.File, filePath string) error {
	if errSave := f.SaveAs(filePath); errSave != nil {
		fmt.Println(errSave, "[GenerateReportRefCode] Error in Save Excel")
		return errSave
	}
	return nil
}