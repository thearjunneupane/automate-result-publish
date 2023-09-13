package demo_entry

import (
	"fmt"
	"log"
	"math/rand"
	"os"
	"path/filepath"
	"time"

	"github.com/tealeg/xlsx"
)

func Demo_entry() {

	if err := os.MkdirAll("publish_result/subjects_marks", os.ModePerm); err != nil {
		log.Fatalf("Error creating results directory: %v", err)
	}
	outputDir := "publish_result/subjects_marks"

	subjects := []string{"PhysicsSubject", "MathSubject", "ChemistrySubject", "ComputerSubject", "BiologySubject"}

	// Initialize a map to keep track of used team names across all subjects
	usedTeamNames := make(map[string]bool)

	// Create a function to generate a unique team name
	generateUniqueTeamName := func() string {
		for {
			// Generate a random team number between 1 and 400
			teamNumber := rand.Intn(400) + 1
			teamName := fmt.Sprintf("CFT - C%d", teamNumber)

			// Check if the team name is already used
			if !usedTeamNames[teamName] {
				usedTeamNames[teamName] = true
				return teamName
			}
		}
	}

	for _, subjectName := range subjects {
		// Create a new Excel file for the current subject
		file := xlsx.NewFile()

		// Create a new sheet in the Excel file
		sheet, err := file.AddSheet("Sheet1")
		if err != nil {
			fmt.Printf("Error creating sheet for %s: %v\n", subjectName, err)
			continue
		}

		// Add headers to the first row (A1 and B1)
		headerRow := sheet.AddRow()
		headerCell1 := headerRow.AddCell()
		headerCell2 := headerRow.AddCell()
		headerCell1.Value = "Team"
		headerCell2.Value = "Marks"

		// Seed the random number generator
		rand.Seed(time.Now().UnixNano())

		// Determine the number of teams based on the subject
		var numTeams int
		if subjectName == "ComputerSubject" || subjectName == "BiologySubject" {
			numTeams = 40
		} else {
			numTeams = 80
		}

		// Generate and add random data to the sheet
		for i := 0; i < numTeams; i++ {
			dataRow := sheet.AddRow()
			teamName := generateUniqueTeamName()
			marks := rand.Intn(20) + 1

			teamCell := dataRow.AddCell()
			marksCell := dataRow.AddCell()

			teamCell.Value = teamName
			marksCell.SetInt(marks)
		}

		// Save the Excel file in the specified directory or the current directory
		fileName := subjectName + ".xlsx"
		outputPath := filepath.Join(outputDir, fileName)
		err = file.Save(outputPath)
		if err != nil {
			fmt.Printf("Error saving Excel file for %s: %v\n", subjectName, err)
			continue
		}

		fmt.Printf("Excel file '%s' created successfully with %d teams in %s.\n", fileName, numTeams, outputPath)
	}
}
