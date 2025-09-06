package main

import (
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Define the list of files to combine in an array (slice).
	// old code
	// filesToCombine := []string{
	// 	"nasabah_baru_1.xlsx",
	// 	"nasabah_baru_2.xlsx",
	// 	// Add more file names here as needed.
	// }

	// Define the list of files to combine in an array (slice).
	// Get the list of files to combine from command-line arguments.
	// The first argument (os.Args[0]) is the program name, so we take the rest.
	// Example usage: go run main.go file1.xlsx file2.xlsx
	filesToCombine := os.Args[1:]

	// If no file names are provided as arguments, print a usage message and exit.
	if len(filesToCombine) == 0 {
		log.Fatal("Usage: go run main.go <file1.xlsx> <file2.xlsx> ...")
	}

	// Create a new Excel file for the combined data.
	combinedFile := excelize.NewFile()
	// Create a stream writer for "Sheet1".
	streamWriter, err := combinedFile.NewStreamWriter("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	totalRowCount := 0

	// Loop through each file in the slice.
	for i, fileName := range filesToCombine {
		log.Printf("Processing %s...", fileName)
		file, err := excelize.OpenFile(fileName)
		if err != nil {
			log.Printf("Error opening %s: %v. Skipping this file.", fileName, err)
			continue
		}

		rows, err := file.Rows("Sheet1")
		if err != nil {
			log.Printf("Error getting rows from %s: %v. Skipping this file.", fileName, err)
			file.Close()
			continue
		}

		// For every file after the first one, skip its header row.
		isHeaderRow := true
		skipHeader := i > 0

		fileRowCount := 0
		for rows.Next() {
			if skipHeader && isHeaderRow {
				isHeaderRow = false
				continue // Skip the header row.
			}
			isHeaderRow = false
			fileRowCount++
			totalRowCount++

			row, err := rows.Columns()
			if err != nil {
				log.Println(err)
				continue
			}

			// Convert []string to []interface{} for SetRow.
			interfaceRow := make([]interface{}, len(row))
			for i, v := range row {
				interfaceRow[i] = v
			}

			// Write the row to the stream writer at the correct position.
			cell, _ := excelize.CoordinatesToCellName(1, totalRowCount)
			if err := streamWriter.SetRow(cell, interfaceRow); err != nil {
				log.Fatal(err)
			}

			if fileRowCount%1000 == 0 { // Log progress every 1000 rows per file.
				log.Printf("Processed %d rows from %s\n", fileRowCount, fileName)
			}
		}

		if err = rows.Close(); err != nil {
			log.Println(err)
		}
		file.Close()
		log.Printf("Finished processing %s. Total rows appended: %d\n", fileName, fileRowCount)
	}

	// Flush the stream writer to write all data to the sheet.
	if err := streamWriter.Flush(); err != nil {
		log.Fatal(err)
	}

	// Save the combined Excel file.
	if err := combinedFile.SaveAs("combined_file.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Excel files combined successfully into combined_file.xlsx! Total rows: %d\n", totalRowCount)
}
