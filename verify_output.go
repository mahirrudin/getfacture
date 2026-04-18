package main

import (
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	f1, err := excelize.OpenFile("result.xlsx")
	if err != nil {
		fmt.Printf("Gagal membuka result.xlsx: %v\n", err)
		os.Exit(1)
	}
	defer f1.Close()

	f2, err := excelize.OpenFile("output.xlsx")
	if err != nil {
		fmt.Printf("Gagal membuka output.xlsx: %v\n", err)
		os.Exit(1)
	}
	defer f2.Close()

	compareSheets(f1, f2, "Summary Faktur")
	compareSheets(f1, f2, "Detail Barang")
}

func compareSheets(f1, f2 *excelize.File, sheetName string) {
	fmt.Printf("\n=== Membandingkan Sheet: %s ===\n", sheetName)
	rows1, err := f1.GetRows(sheetName)
	if err != nil {
		fmt.Printf("Error get rows %s from result.xlsx: %v\n", sheetName, err)
		return
	}
	rows2, err := f2.GetRows(sheetName)
	if err != nil {
		fmt.Printf("Error get rows %s from output.xlsx: %v\n", sheetName, err)
		return
	}

	if len(rows1) != len(rows2) {
		fmt.Printf("PERBEDAAN JUMLAH BARIS: result.xlsx punya %d, output.xlsx punya %d\n", len(rows1), len(rows2))
	}

	minRows := len(rows1)
	if len(rows2) < minRows {
		minRows = len(rows2)
	}

	for i := 0; i < minRows; i++ {
		r1 := rows1[i]
		r2 := rows2[i]

		minCols := len(r1)
		if len(r2) < len(r1) { // allow output to have more or less cols, but check overlapping
			minCols = len(r2)
		} else if len(r1) < len(r2) {
			// pad r1 if needed, but realistically we only compare up to what result has
			minCols = len(r1)
		}

		for j := 0; j < minCols; j++ {
			v1 := strings.TrimSpace(r1[j])
			var v2 string
			if j < len(r2) {
				v2 = strings.TrimSpace(r2[j])
			}

			// if same strings, skip
			if v1 == v2 {
				continue
			}

			// try numeric comparison
			f1, err1 := strconv.ParseFloat(v1, 64)
			f2, err2 := strconv.ParseFloat(v2, 64)
			if err1 == nil && err2 == nil {
				if math.Abs(f1-f2) > 0.01 {
					fmt.Printf("Beda di Baris %d Kolom %d:\n  Expected (result): %s\n  Got (output): %s\n", i+1, j+1, v1, v2)
				}
			} else {
				// not both numeric
				if v1 != "" || v2 != "" {
					fmt.Printf("Beda di Baris %d Kolom %d:\n  Expected (result): '%s'\n  Got (output): '%s'\n", i+1, j+1, v1, v2)
				}
			}
		}
	}
}
