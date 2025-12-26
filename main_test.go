package main

import (
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

func tempExcel(t *testing.T) string {
	t.Helper()

	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Report Title")
	f.SetCellValue("Sheet1", "A3", "Colleague ID")
	f.SetCellValue("Sheet1", "B3", "Preferred First Name")
	f.SetCellValue("Sheet1", "C3", "Legal Last Name")
	f.SetCellValue("Sheet1", "D3", "Manager Name")

	f.SetCellValue("Sheet1", "A4", "100")
	f.SetCellValue("Sheet1", "B4", "Jane")
	f.SetCellValue("Sheet1", "C4", "Doe")
	f.SetCellValue("Sheet1", "D4", "Boss One")

	tmp, _ := os.CreateTemp("", "*.xlsx")
	tmp.Close()
	f.SaveAs(tmp.Name())
	return tmp.Name()
}

func TestBuildSupervisorMap(t *testing.T) {
	file := tempExcel(t)
	defer os.Remove(file)

	m, err := buildSupervisorMap(file, "Sheet1")
	if err != nil {
		t.Fatal(err)
	}

	info, ok := m["100"]
	if !ok {
		t.Fatal("expected ID 100")
	}

	if info.Name != "Jane Doe" {
		t.Fatalf("expected Jane Doe, got %s", info.Name)
	}
}

func TestUpdateVLookup(t *testing.T) {
	f := excelize.NewFile()
	f.NewSheet("vlookup")
	f.SetCellValue("vlookup", "A1", "Driver Name")
	f.SetCellValue("vlookup", "B1", "Driver Number")
	f.SetCellValue("vlookup", "C1", "Supervisor")

	lookup := map[string]SupervisorInfo{
		"100": {Name: "Jane Doe", Supervisor: "Boss One"},
	}

	if err := updateVLookupSheet(f, "vlookup", lookup, false); err != nil {
		t.Fatal(err)
	}

	val, _ := f.GetCellValue("vlookup", "A2")
	if val != "Jane Doe" {
		t.Fatalf("expected Jane Doe, got %s", val)
	}
}
