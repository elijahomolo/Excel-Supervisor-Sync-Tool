package main

import (
	"flag"
	"fmt"
	"log"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

type SupervisorInfo struct {
	Name       string
	Supervisor string
}

func main() {
	masterFile := flag.String("master", "", "ALLOPS master Excel file")
	targetFile := flag.String("target", "", "Shorts report Excel file")
	masterSheet := flag.String("master-sheet", "Sheet1", "Master sheet name")
	vlookupSheet := flag.String("vlookup-sheet", "vlookup", "VLOOKUP sheet")
	outFile := flag.String("out", "output.xlsx", "Output file")
	dryRun := flag.Bool("dry-run", false, "Dry run (no writes)")
	flag.Parse()

	if *masterFile == "" || *targetFile == "" {
		log.Fatal("master and target files are required")
	}

	lookup, err := buildSupervisorMap(*masterFile, *masterSheet)
	if err != nil {
		log.Fatal(err)
	}

	f, err := excelize.OpenFile(*targetFile)
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	if err := updateVLookupSheet(f, *vlookupSheet, lookup, *dryRun); err != nil {
		log.Fatal(err)
	}

	if *dryRun {
		fmt.Println("Dry run enabled — no file written.")
		return
	}

	if err := f.SaveAs(*outFile); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Update completed successfully.")
}

/* ---------------- Header Matching ---------------- */

var nonAlphaNum = regexp.MustCompile(`[^a-z0-9]+`)

func normalizeHeader(s string) string {
	s = strings.ReplaceAll(s, "\u00A0", " ")
	s = strings.ToLower(strings.TrimSpace(s))
	return nonAlphaNum.ReplaceAllString(s, "")
}

func fuzzyMatch(have, want string) bool {
	h := normalizeHeader(have)
	w := normalizeHeader(want)
	return h == w || strings.Contains(h, w) || strings.Contains(w, h)
}

func findHeaderRow(rows [][]string, required []string) (int, []int) {
	for i, row := range rows {
		matches := 0
		for _, cell := range row {
			for _, r := range required {
				if fuzzyMatch(cell, r) {
					matches++
				}
			}
		}
		if matches >= len(required)-1 {
			return i, findColumns(row, required...)
		}
	}
	log.Fatalf("Could not locate header row with required columns: %v", required)
	return -1, nil
}

func findColumns(headers []string, wanted ...string) []int {
	indexes := make([]int, len(wanted))
	for i := range indexes {
		indexes[i] = -1
	}

	for i, header := range headers {
		for j, want := range wanted {
			if indexes[j] == -1 && fuzzyMatch(header, want) {
				indexes[j] = i
			}
		}
	}

	for i, idx := range indexes {
		if idx == -1 {
			log.Fatalf("Missing required column: %s", wanted[i])
		}
	}
	return indexes
}

/* ---------------- Master Parsing ---------------- */

func buildSupervisorMap(filename, sheet string) (map[string]SupervisorInfo, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	headerRow, cols := findHeaderRow(
		rows,
		[]string{
			"Colleague ID",
			"Preferred First Name",
			"Legal Last Name",
			"Manager Name",
		},
	)

	result := make(map[string]SupervisorInfo)
	for i := headerRow + 1; i < len(rows); i++ {
		row := rows[i]
		id := value(row, cols[0])
		if id == "" {
			continue
		}

		name := strings.TrimSpace(
			value(row, cols[1]) + " " + value(row, cols[2]),
		)

		result[id] = SupervisorInfo{
			Name:       name,
			Supervisor: value(row, cols[3]),
		}
	}
	return result, nil
}

/* ---------------- VLOOKUP Update ---------------- */

func updateVLookupSheet(
	f *excelize.File,
	sheet string,
	lookup map[string]SupervisorInfo,
	dryRun bool,
) error {

	rows, err := f.GetRows(sheet)
	if err != nil {
		return err
	}

	cols := findColumns(
		rows[0],
		"Driver Name",
		"Driver Number",
		"Supervisor",
	)

	existing := map[string]int{}
	for i := 1; i < len(rows); i++ {
		id := value(rows[i], cols[1])
		if id != "" {
			existing[id] = i + 1
		}
	}

	nextRow := len(rows) + 1

	for id, info := range lookup {
		if rowNum, ok := existing[id]; ok {
			if !dryRun {
				f.SetCellValue(sheet, cell(cols[2], rowNum), info.Supervisor)
			}
		} else {
			if !dryRun {
				f.SetCellValue(sheet, cell(cols[0], nextRow), info.Name)
				f.SetCellValue(sheet, cell(cols[1], nextRow), id)
				f.SetCellValue(sheet, cell(cols[2], nextRow), info.Supervisor)
			}
			nextRow++
		}
	}
	return nil
}

/* ---------------- Helpers ---------------- */

func value(row []string, idx int) string {
	if idx >= len(row) {
		return ""
	}
	return strings.TrimSpace(row[idx])
}

func cell(col, row int) string {
	c, _ := excelize.CoordinatesToCellName(col+1, row)
	return c
}
