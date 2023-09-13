# Automate-Result-Publish

A <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/0/05/Go_Logo_Blue.svg/1280px-Go_Logo_Blue.svg.png" height="18"> program that automates the processing of Excel data by sorting and organizing it into new Excel files, and serving into the server.

## Description

This program converts Excel spreadsheet data into HTML tables and serves it via a web application.

<details>
  <summary><b>automate-result-publish-static</b></summary>
  <ul>
    <li>Generate two excel file with marks and without marks.</li>
    <li>Generate a new html file from that excel file.</li>
  </ul>
</details>

<details>
  <summary><b>automate-result-publish-dynamic</b></summary>
  <ul>
    <li>Generate two excel file with marks and without marks.</li>
    <li>Parse the sorted excel file into a html template.</li>
    <li>Publish a web server at <i>localhost:8080</i></li>
  </ul>
</details>

## Getting Started

### Dependencies

* Go
* Make -> ```choco install make```
* [xlsx](https://github.com/tealeg/xlsx)

### Installing

* ```git clone https://github.com/thearjunneupane/automate-result-publish.git```
* ```cd automate-result-publish```

### Executing program

* ```cd automate-result-publish-static``` or ```cd automate-result-publish-dynamic```
* ```make run```

### Directory
```

automate-result-publish 
│   
├── automate-result-publish-dynamic 
│   ├── demo_entry  
│   │   └── demo_entry.go             # This file generate a demo excel file with Teams and Marks   
│   ├── publish_result  
│   │    ├── subjects_marks           # This folder will be created by demo_entry.go    
│   │    │   └── ...                  # Here the 5 demo excel files with Team name and Marks is created.    
│   │    ├── results                  # This folder will be created by demo_entry.go    
│   │    │   └── ...                  # Here the 2 result excel files withmarks and withnamesonly is created.   
│   │    └── publish.go               # This file generate the result by sorting the selected teams 
│   └── web 
│        ├── static 
│        │   ├── index.css   
│        │   └── result.css  
│        ├── templates  
│        │   ├── index.html  
│        │   └── result.html 
│        ├── main.go    
│        ├── Makefile     
│        └── ...                      # Here the 2 result excel files withmarks and withnamesonly is created. 
│   
├── automate-result-publish-static  
│   ├── demo_entry  
│   │   └── demo_entry.go   
│   ├── publish_result  
│   │    ├── subjects_marks 
│   │    │   └── ...    
│   │    ├── results    
│   │    │   └── ...    
│   │    └── publish.go               # This file generate the result by sorting the selected teams in result.html file.    
│   ├── static  
│   │    ├──result.html               # This file will be created or overwritten by publish.go  
│   │    └──result.css  
│   ├── main.go     
│   └── Makefile    
│   
├── go.mod  
└── go.sum

```