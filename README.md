# Test Evidence Creator

Test Evidence Creator is a Go-based utility tool for generating unit test evidence by inserting images into Excel sheets. 
It automates the process of organizing test results in a structured manner within an Excel template, making it easier to review and share evidence.

## Features

- Walks through a specified folder to fetch and sort image files.
- Inserts images horizontally into a specified sheet in an Excel template.
- Supports scaling images to a desired size.
- Allows insertion of page breaks after each image.
- Handles popular image formats such as PNG and JPEG.

## Requirements

- [Go](https://go.dev/) (1.20 or later recommended)
- Excel template file
- Folder containing images

## Installation

1. Clone the repository

```bash
   git clone https://github.com/your-username/TestEvidenceCreator.git
```

2. Navigate to the project directory

```bash
   cd TestEvidenceCreator
```

3. Install dependencies
```bash
   go mod tidy
```

##Usage

Run the application with the following flags

```bash
go run main.go -f="./Evidences" -e="sample.xlsx"
```

##Output

The tool will:

- Detect all sub folders from provided parent folder.
- Sort the sub folders and add these sorted sub folders as sheet.
- Sheet names will be same as (# + sub folder names).
- Insert images starting from cell B4 in the specified sheet.
- Scale the images to fit within the desired dimensions.
- Insert page breaks after each images.
- Sheet name #0 will be rename as #Preparation sheet.
- PageLayout will be A4, Landscape orientation and "Fit All Rows on One Page".
- Sheets in the Excel file : [Final Template #Preparation #1 #2 #3]


