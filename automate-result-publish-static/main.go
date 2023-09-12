package main

import (
	"html/template"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

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

func main() {
	fileName := "resultwithmarks.xlsx" // Replace with the actual filename
	htmlTable, err := excelToHTML(fileName)
	if err != nil {
		panic(err)
	}

	err = generateHTMLFile(htmlTable)
	if err != nil {
		panic(err)
	}

	println("HTML file generated successfully: result.html")
}

func generateHTMLFile(htmlTable string) error {
	// Define the HTML template
	templateStr := `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Result Page</title>
    <link rel="stylesheet" type="text/css" href="result.css">
</head>
<body>
    <div class="container">
        <header>
            <h1>Result Page</h1>
        </header>
        <div class="description"><b>Slide if all subject marks are not seen.</b></div>
        <div class="table-container">
            {{.HTMLTable | safeHTML}}
        </div>
    </div>
</body>
</html>`

	// Create a template from the template string
	tmpl, err := template.New("htmlTemplate").Funcs(template.FuncMap{
		"safeHTML": func(s string) template.HTML { return template.HTML(s) },
	}).Parse(templateStr)
	if err != nil {
		return err
	}

	// Create or overwrite the HTML file
	file, err := os.Create("static/result.html")
	if err != nil {
		return err
	}
	defer file.Close()

	// Execute the template and write the HTML to the file
	data := map[string]interface{}{
		"HTMLTable": htmlTable, // Use template.HTML to prevent escaping HTML
	}
	if err := tmpl.Execute(file, data); err != nil {
		return err
	}

	return nil
}
