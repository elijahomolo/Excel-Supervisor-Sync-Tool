package main

import (
	"flag"
	"fmt"
	"log"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

/* ===================== Types ===================== */

type SupervisorInfo struct {
	Name       string
	Supervisor string
}

type ValidationReport struct {
	MissingInAllops []string
}

/* ===================== Main ===================== */

func main() {
	masterFile := flag.String("master", "", "ALLOPS master Excel file")
	targetFile := flag.String("target", "", "Shorts report Excel file")

	masterSheet := flag.String("master-sheet", "Sheet1", "ALLOPS sheet")
	dataSheet := flag.String("data-sheet", "Data", "Shorts Data sheet")
	vlookupSheet := flag.String("vlookup-sheet", "vlookup", "Shorts vlookup sheet")

	outFile := flag.String("out", "output.xlsx", "Output file")
	dryRun := flag.Bool("dry-run", false, "Dry run")
	flag.Parse()

	if *masterFile == "" || *targetFile == "" {
		log.Fatal("master and target files are required")
	}

	allops, err := buildSupervisorMap(*masterFile, *masterSheet)
	if err != nil {
		log.Fatal(err)
	}

	shorts, err := excelize.OpenFile(*targetFile)
	if err != nil {
		log.Fatal(err)
	}
	defer shorts.Close()

	driverIDs, err := readDriverIDsFromData(shorts, *dataSheet)
	if err != nil {
		log.Fatal(err)
	}

	if err := cleanupVLookupDuplicates(
		shorts,
		*vlookupSheet,
		*dryRun,
	); err != nil {
		log.Fatal(err)
	}

	report, err := updateVLookup(
		shorts,
		*vlookupSheet,
		driverIDs,
		allops,
		*dryRun,
	)
	if err != nil {
		log.Fatal(err)
	}

	if len(report.MissingInAllops) > 0 {
		fmt.Println("WARNING: Drivers missing in ALLOPS:")
		for _, id := range report.MissingInAllops {
			fmt.Printf("  - %s\n", id)
		}
	}

	if *dryRun {
		fmt.Println("Dry run enabled — no file written.")
		return
	}

	if err := shorts.SaveAs(*outFile); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Update completed successfully.")
}

/* ===================== Header Matching ===================== */

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

/* ===================== ALLOPS Parsing ===================== */

func buildSupervisorMap(file, sheet string) (map[string]SupervisorInfo, error) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	cols := findColumns(
		rows[0],
		"Colleague ID",
		"Preferred First Name",
		"Legal Last Name",
		"Manager Name",
	)

	result := map[string]SupervisorInfo{}

	for i := 1; i < len(rows); i++ {
		id := value(rows[i], cols[0])
		if id == "" {
			continue
		}

		result[id] = SupervisorInfo{
			Name: strings.TrimSpace(
				value(rows[i], cols[1]) + " " + value(rows[i], cols[2]),
			),
			Supervisor: value(rows[i], cols[3]),
		}
	}

	return result, nil
}

/* ===================== Data Sheet Parsing ===================== */

func readDriverIDsFromData(
	f *excelize.File,
	sheet string,
) ([]string, error) {

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	cols := findColumns(rows[0], "Driver ID")

	seen := map[string]bool{}
	var ids []string

	for i := 1; i < len(rows); i++ {
		id := value(rows[i], cols[0])
		if id == "" {
			continue
		}

		// IGNORE duplicates, do not fail
		if seen[id] {
			log.Printf(
				"duplicate Driver ID ignored in Data sheet: %s (row %d)",
				id,
				i+1,
			)
			continue
		}

		seen[id] = true
		ids = append(ids, id)
	}

	return ids, nil
}

/* ===================== vlookup Update ===================== */
func cleanupVLookupDuplicates(
	f *excelize.File,
	sheet string,
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
	)

	type seenEntry struct {
		name string
		row  int
	}

	seen := map[string]seenEntry{}
	var rowsToDelete []int

	for i := 1; i < len(rows); i++ {
		rowNum := i + 1
		id := value(rows[i], cols[1])
		name := value(rows[i], cols[0])

		if id == "" {
			continue
		}

		if prev, exists := seen[id]; exists {
			// Same ID + same name → safe duplicate
			if prev.name == name {
				rowsToDelete = append(rowsToDelete, rowNum)
				continue
			}

			// Same ID + different name → corruption
			return fmt.Errorf(
				"conflicting Driver Name for Driver ID %s: '%s' vs '%s'",
				id,
				prev.name,
				name,
			)
		}

		seen[id] = seenEntry{name: name, row: rowNum}
	}

	// Delete bottom-up to preserve row indexes
	for i := len(rowsToDelete) - 1; i >= 0; i-- {
		if !dryRun {
			f.RemoveRow(sheet, rowsToDelete[i])
		}
	}

	if len(rowsToDelete) > 0 {
		log.Printf(
			"vlookup cleanup: removed %d duplicate rows",
			len(rowsToDelete),
		)
	}

	return nil
}

func updateVLookup(
	f *excelize.File,
	sheet string,
	driverIDs []string,
	allops map[string]SupervisorInfo,
	dryRun bool,
) (*ValidationReport, error) {

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
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
		if id == "" {
			continue
		}
		if _, dup := existing[id]; dup {
			return nil, fmt.Errorf(
				"duplicate Driver ID found in vlookup sheet: %s",
				id,
			)
		}
		existing[id] = i + 1
	}

	nextRow := len(rows) + 1
	report := &ValidationReport{}

	for _, id := range driverIDs {
		info, ok := allops[id]
		if !ok {
			report.MissingInAllops = append(report.MissingInAllops, id)
			continue
		}

		if row, exists := existing[id]; exists {
			if !dryRun {
				f.SetCellValue(sheet, cell(cols[2], row), info.Supervisor)
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

	return report, nil
}

/* ===================== Helpers ===================== */

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
