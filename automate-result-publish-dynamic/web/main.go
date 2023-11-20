package main

import (
	"fmt"
	"html/template"
	"log"
	"net/http"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
	"github.com/thearjunneupane/automate-result-publish/automate-result-publish-dynamic/demo_entry"
	"github.com/thearjunneupane/automate-result-publish/automate-result-publish-dynamic/publish_result"
)

// mergeCellsin a given range of rows and columns.
func mergeCells(sheet *xlsx.Sheet, startRow, endRow, startCol, endCol int) {
	for i := startRow; i <= endRow; i++ {
		for j := startCol; j <= endCol; j++ {
			cell := sheet.Cell(i, j)
			cell.HMerge = endCol - startCol
			cell.VMerge = endRow - startRow
		}
	}
}

// excelToHTML converts Excel data to an HTML table.
func excelToHTML(fileName string) (string, error) {
	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		return "", err
	}

	// Assuming the first sheet in the Excel file
	sheet := file.Sheets[0]

	// Create an HTML table
	var htmlTable strings.Builder
	htmlTable.WriteString("<table>")
	for rowIndex, row := range sheet.Rows {
		htmlTable.WriteString("<tr>")
		for cellIndex, cell := range row.Cells {
			// Use <th> for the first row (headers), <td> for others (data)
			cellType := "td"
			if rowIndex == 0 {
				cellType = "th"
			}
			// Create the cell with the appropriate type
			if cellType == "td" {
				if rowIndex == 1 {
					if cellIndex%2 == 0 {
						htmlTable.WriteString("<" + cellType + " class=\"team-name\" " + ">" + cell.Value + "</" + cellType + ">")
						continue
					} else {
						htmlTable.WriteString("<" + cellType + " class=\"team-marks\" " + ">" + cell.Value + "</" + cellType + ">")
						continue
					}
				}
			} else if cellType == "th" {
				// For Showing Excel File with Marks only the -1 should be removed
				if cellIndex%2 == 0 && cellIndex != len(row.Cells) && fileName == "resultwithoutmarks.xlsx" {
					htmlTable.WriteString("<" + cellType + " class=\"subject-name\" colspan=\"2\" " + ">" + cell.Value + "</" + cellType + ">")
					continue
				} else if cellIndex%2 == 0 && cellIndex != len(row.Cells)-1 && fileName == "resultwithmarks.xlsx" {
					htmlTable.WriteString("<" + cellType + " class=\"subject-name\" colspan=\"2\" " + ">" + cell.Value + "</" + cellType + ">")
					continue
				} else {
					continue
				}
			}
			htmlTable.WriteString("<" + cellType + ">" + cell.Value + "</" + cellType + ">")
		}
		// if rowIndex != 0 {
		// 	htmlTable.WriteString("<td></td>")
		// }
		htmlTable.WriteString("</tr>")
	}
	htmlTable.WriteString("</table>")

	return htmlTable.String(), nil
}

func resultHandler(w http.ResponseWriter, r *http.Request) {
	// Generate HTML from Excel data
	htmlTable, err := excelToHTML("resultwithmarks.xlsx")
	if err != nil {
		http.Error(w, "Error generating HTML: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// Read and parse HTML template
	tmpl, err := template.ParseFiles("automate-result-publish-dynamic/web/templates/result.html")
	if err != nil {
		http.Error(w, "Error parsing HTML template: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// Define data to be passed to the template
	data := struct {
		HTMLTable template.HTML
	}{
		HTMLTable: template.HTML(htmlTable),
	}

	// Execute the template with data and write to response
	if err := tmpl.Execute(w, data); err != nil {
		http.Error(w, "Error executing template: "+err.Error(), http.StatusInternalServerError)
	}
}

func main() {
	demo_entry.Demo_entry()
	publish_result.Publish()
	// Define a route for displaying the result HTML
	http.HandleFunc("/result", resultHandler)
	http.HandleFunc("/", indexHandler)

	// Serve static files (CSS)
	fs := http.FileServer(http.Dir("static/"))
	http.Handle("/static/", http.StripPrefix("/static/", fs))

	// Start the HTTP server
	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}
	fmt.Printf("Server is listening on :%s...\n", port)
	log.Fatal(http.ListenAndServe(fmt.Sprintf(":%s", port), nil))
}

func indexHandler(w http.ResponseWriter, r *http.Request) {
	tmpl, err := template.ParseFiles("automate-result-publish-dynamic/web/templates/index.html")
	if err != nil {
		http.Error(w, "Error parsing HTML template: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// Execute the template with data and write to response
	if err := tmpl.Execute(w, nil); err != nil {
		http.Error(w, "Error executing template: "+err.Error(), http.StatusInternalServerError)
	}
}
