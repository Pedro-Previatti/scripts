package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/tealeg/xlsx"
)


func xlsxToCsv(inputFile string, outputFile string) error {
	wb, err := xlsx.OpenFile(inputFile)
	if err != nil {
		return fmt.Errorf("could not open XLSX file: %v", err)
	}

	csvFile, err := os.Create(outputFile)
	if err != nil {
		return fmt.Errorf("could not create CSV file: %v", err)
	}
	defer csvFile.Close()

	writer := csv.NewWriter(csvFile)
	defer writer.Flush()

	for _, sheet := range wb.Sheets {
		for _, row := range sheet.Rows {
			var rowData []string
			for _, cell := range row.Cells {
				text := cell.String()
				rowData = append(rowData, text)
			}
			writer.Write(rowData)
		}
	}

	return nil
}

func csvToXlsx(inputFile string, outputFile string) error {
	csvFile, err := os.Open(inputFile)
	if err != nil {
		return fmt.Errorf("could not open CSV file: %v", err)
	}
	defer csvFile.Close()

	reader := csv.NewReader(csvFile)

	wb := xlsx.NewFile()
	sheet, err := wb.AddSheet("Sheet1")
	if err != nil {
		return fmt.Errorf("could not create XLSX sheet: %v", err)
	}

	for {
		record, err := reader.Read()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			return fmt.Errorf("could not read CSV file: %v", err)
		}

		row := sheet.AddRow()
		for _, value := range record {
			cell := row.AddCell()
			cell.Value = value
		}
	}

	err = wb.Save(outputFile)
	if err != nil {
		return fmt.Errorf("could not save XLSX file: %v", err)
	}

	return nil
}

func main() {
	formatFlag := flag.String("f", "", "The format to convert to: 'csv' or 'xlsx'")
	outputFlag := flag.String("o", "", "The output filename")
	flag.Parse()

	args := flag.Args()
	if len(args) < 1 {
		log.Fatal("Usage: go run convert.go -f <format> -o <output_filename> <input_filename>")
	}
	inputFile := args[0]

	if *formatFlag == "" || *outputFlag == "" || inputFile == "" {
		log.Fatal("Usage: go run convert.go -f <format> -o <output_filename> <input_filename>")
	}

	switch *formatFlag {
	case "csv":
		err := xlsxToCsv(inputFile, *outputFlag)
		if err != nil {
			log.Fatalf("Error converting XLSX to CSV: %v", err)
		}
		fmt.Println("Conversion to CSV completed successfully.")
	case "xlsx":
		err := csvToXlsx(inputFile, *outputFlag)
		if err != nil {
			log.Fatalf("Error converting CSV to XLSX: %v", err)
		}
		fmt.Println("Conversion to XLSX completed successfully.")
	default:
		log.Fatal("Unsupported format. Use 'csv' or 'xlsx'.")
	}
}
