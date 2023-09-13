package publish_result

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strconv"

	"github.com/tealeg/xlsx"
)

// mergeCells merges cells in a given range of rows and columns.
func mergeCells(sheet *xlsx.Sheet, startRow, endRow, startCol, endCol int) {
	for i := startRow; i <= endRow; i++ {
		for j := startCol; j <= endCol; j++ {
			cell := sheet.Cell(i, j)
			cell.HMerge = endCol - startCol
			cell.VMerge = endRow - startRow
		}
	}
}

// processSubjectWithMarks processes a subject's Excel file and appends the result to the output sheet (with marks).
func processSubjectWithMarks(outputSheet *xlsx.Sheet, subjectName string, columnIndex, numTeams int) {
	// Open the Excel file for the subject with name subjectName from "subjects_marks" dir.
	fileName := filepath.Join("../publish_result/subjects_marks", subjectName+".xlsx")
	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		log.Fatalf("Error opening file %s: %v", fileName, err)
	}

	// Get the first sheet (assuming there's only one sheet).
	sheet := file.Sheets[0]

	// Create a slice to store rows.
	rows := make([]xlsx.Row, len(sheet.Rows))
	for i := range sheet.Rows {
		rows[i] = *sheet.Rows[i]
	}

	// Sort the rows based on the "Marks" column in descending order and then by team names.
	sort.SliceStable(rows[1:], func(i, j int) bool {
		marksA, _ := strconv.Atoi(rows[i+1].Cells[1].Value)
		marksB, _ := strconv.Atoi(rows[j+1].Cells[1].Value)
		teamA := rows[i+1].Cells[0].Value
		teamB := rows[j+1].Cells[0].Value

		if marksA != marksB {
			return marksA > marksB
		}

		return teamA < teamB
	})

	// Get marks of 40th team.
	numTeamsMarks, _ := strconv.Atoi(rows[numTeams].Cells[1].Value)

	// Select the top 40 teams and teams with equal marks to the 40th team.
	selectedTeams := make([][]interface{}, 0)
	selectedTeams = append(selectedTeams, []interface{}{"Team", "Marks"})

	for i := 1; i < len(rows); i++ {
		row := rows[i]
		marks, _ := strconv.Atoi(row.Cells[1].Value)
		team := row.Cells[0].Value

		if marks >= numTeamsMarks {
			selectedTeams = append(selectedTeams, []interface{}{team, marks})
		}
	}

	// Merge cells of subject name header.
	mergeCells(outputSheet, 0, 0, columnIndex, columnIndex+1)

	// Set the subject name header in the merged cell.
	cell := outputSheet.Cell(0, columnIndex)
	cell.Value = subjectName

	// Set headers "Team" and "Marks" for this subject.
	outputSheet.Cell(1, columnIndex).Value = "Team"
	outputSheet.Cell(1, columnIndex+1).Value = "Marks"

	// Add selected teams' data for this subject.
	for i, rowData := range selectedTeams[1:] {
		team := rowData[0].(string)
		marks := rowData[1].(int) // Cast to integer
		outputSheet.Cell(i+2, columnIndex).Value = team
		outputSheet.Cell(i+2, columnIndex+1).SetInt(marks) // Set as integer
	}
}

// processSubjectNamesOnly processes a subject's Excel file and appends the result to the output sheet with only team names.
func processSubjectNamesOnly(outputSheet *xlsx.Sheet, subjectName string, columnIndex, numTeams int) {
	// Open the Excel file for the subject you want to process.
	fileName := filepath.Join("../publish_result/subjects_marks", subjectName+".xlsx")
	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		log.Fatalf("Error opening file %s: %v", fileName, err)
	}

	// Get the first sheet (assuming there's only one sheet).
	sheet := file.Sheets[0]

	// Create a slice to store rows.
	rows := make([]xlsx.Row, len(sheet.Rows))
	for i := range sheet.Rows {
		rows[i] = *sheet.Rows[i]
	}

	// Sort the rows based on the "Marks" column in descending order and then by team names.
	sort.SliceStable(rows[1:], func(i, j int) bool {
		marksA, _ := strconv.Atoi(rows[i+1].Cells[1].Value)
		marksB, _ := strconv.Atoi(rows[j+1].Cells[1].Value)
		teamA := rows[i+1].Cells[0].Value
		teamB := rows[j+1].Cells[0].Value

		if marksA != marksB {
			return marksA > marksB
		}

		return teamA < teamB
	})

	// Find the marks of the 40th team.
	numTeamsMarks, _ := strconv.Atoi(rows[numTeams].Cells[1].Value)

	// Select the top 40 teams and teams with equal marks to the 40th team.
	selectedTeams := make([]string, 0)
	selectedTeams = append(selectedTeams, "Team") // Add "Team" header

	for i := 1; i < len(rows); i++ {
		row := rows[i]
		marks, _ := strconv.Atoi(row.Cells[1].Value)
		team := row.Cells[0].Value

		if marks >= numTeamsMarks {
			selectedTeams = append(selectedTeams, team)
		}
	}

	// Merge cells for the subject name header.
	mergeCells(outputSheet, 0, 0, columnIndex, columnIndex)

	// Set the subject name header in the merged cell.
	cell := outputSheet.Cell(0, columnIndex)
	cell.Value = subjectName

	// Add selected teams' data for this subject. (name only)
	for i, team := range selectedTeams {
		outputSheet.Cell(i+1, columnIndex).Value = team
	}
}

func Publish() {
	resultsDir := "../publish_result/results"
	altresultsDir := ""
	if err := os.MkdirAll(resultsDir, os.ModePerm); err != nil {
		log.Fatalf("Error creating results directory: %v", err)
	}

	// Create a new Excel file to store the selected teams with marks.
	outputFileWithMarks := xlsx.NewFile()

	// Create a new Excel file to store the selected teams without marks.
	outputFileWithoutMarks := xlsx.NewFile()

	// Add a sheet for the subjects in all output files.
	outputSheetWithMarks, err := outputFileWithMarks.AddSheet("SelectedTeamsWithMarks")
	if err != nil {
		log.Fatalf("Error adding sheet: %v", err)
	}
	outputSheetWithoutMarks, err := outputFileWithoutMarks.AddSheet("SelectedTeamsWithoutMarks")
	if err != nil {
		log.Fatalf("Error adding sheet: %v", err)
	}
	mainResultWithMarks, err := outputFileWithMarks.AddSheet("resultwithmarks")
	if err != nil {
		log.Fatalf("Error adding sheet: %v", err)
	}
	mainResultWithoutMarks, err := outputFileWithoutMarks.AddSheet("resultwithoutmarks")
	if err != nil {
		log.Fatalf("Error adding sheet: %v", err)
	}

	// Define the subject names.
	subjectNames := []string{"PhysicsSubject", "ComputerSubject", "MathSubject", "ChemistrySubject", "BiologySubject"}

	// Process each subject and append the result to the output sheets.
	columnIndex := 0
	for _, subjectName := range subjectNames {
		numTeams := 40 // (Number of teams to select)
		if subjectName == "ComputerSubject" || subjectName == "BiologySubject" {
			numTeams = 20 // ( Number of teams to select for ComputerSubject and BiologySubject )
		}

		processSubjectWithMarks(outputSheetWithMarks, subjectName, columnIndex, numTeams)
		processSubjectNamesOnly(outputSheetWithoutMarks, subjectName, columnIndex, numTeams)
		processSubjectWithMarks(mainResultWithMarks, subjectName, columnIndex, numTeams)
		processSubjectNamesOnly(mainResultWithoutMarks, subjectName, columnIndex, numTeams)

		columnIndex += 2 // Move to the next column for the next subject.
	}

	// Save the output files.
	outputFileNameWithMarks := filepath.Join(resultsDir, "SelectedTeamsWithMarks.xlsx")
	if err := outputFileWithMarks.Save(outputFileNameWithMarks); err != nil {
		log.Fatalf("Error saving output file with marks: %v", err)
	}
	fmt.Printf("Selected teams with marks saved to %s\n", outputFileNameWithMarks)

	outputFileNameWithoutMarks := filepath.Join(resultsDir, "SelectedTeamsWithoutMarks.xlsx")
	if err := outputFileWithoutMarks.Save(outputFileNameWithoutMarks); err != nil {
		log.Fatalf("Error saving output file without marks: %v", err)
	}

	fmt.Printf("Selected teams without marks saved to %s\n", outputFileNameWithoutMarks)

	mainResultWithMarksFileName := filepath.Join(altresultsDir, "resultwithmarks.xlsx")
	if err := outputFileWithMarks.Save(mainResultWithMarksFileName); err != nil {
		log.Fatalf("Error saving output file without marks: %v", err)
	}

	fmt.Printf("Selected teams with marks saved to %s\n", mainResultWithMarksFileName)

	mainResultWithoutMarksFileName := filepath.Join(altresultsDir, "resultwithoutmarks.xlsx")
	if err := outputFileWithoutMarks.Save(mainResultWithoutMarksFileName); err != nil {
		log.Fatalf("Error saving output file without marks: %v", err)
	}

	fmt.Printf("Selected teams without marks saved to %s\n", mainResultWithoutMarksFileName)

}
