package main

import (
	"flag"
	"fmt"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"os"
	"path/filepath"
	"regexp"
	"sort"

	"github.com/xuri/excelize/v2"
)

// ImageInfo holds image file details
type ImageInfo struct {
	FilePath string
}

func main() {
	// Define flags for the image folder path and sheet name
	folderPath := flag.String("folder", "", "Path to the folder containing images")
	sheetName := flag.String("sheet", "", "Name of the sheet")
	templatePath := flag.String("excel", "", "Name of the excel")

	// Parse the command-line flags
	flag.Parse()

	// Validate inputs
	if err := validateInputs(*folderPath, *sheetName, *templatePath); err != nil {
		fmt.Println(err)
		return
	}

	// Get sorted image files
	imageFiles, err := getImageFiles(*folderPath)
	if err != nil {
		fmt.Printf("Error walking through the folder: %v\n", err)
		return
	}

	// Open the existing Excel template file
	f, err := openExcelFile(*templatePath)
	if err != nil {
		fmt.Printf("Failed to open template file: %v\n", err)
		return
	}

	// Start inserting images at a specific row and column
	startCell := "B4" // Starting position for the images
	err = pasteImagesHorizontally(f, *sheetName, imageFiles, startCell)
	if err != nil {
		fmt.Printf("Error inserting images: %v\n", err)
		return
	}

	// Save the changes directly to the same file
	if err := saveExcelFile(f); err != nil {
		fmt.Printf("Failed to save updated file: %v\n", err)
		return
	}

	fmt.Println("Images inserted successfully into the template file:", *templatePath)
}

// validateInputs checks if the provided folder, sheet, and excel file paths are valid.
func validateInputs(folderPath, sheetName, templatePath string) error {
	if folderPath == "" {
		return fmt.Errorf("Please provide the image folder path using the -folder flag.")
	}
	if sheetName == "" {
		return fmt.Errorf("Please provide the sheet name using the -sheet flag.")
	}
	if templatePath == "" {
		return fmt.Errorf("Please provide the excel file path using the -excel flag.")
	}
	if _, err := os.Stat(folderPath); os.IsNotExist(err) {
		return fmt.Errorf("The folder path does not exist: %s", folderPath)
	}
	return nil
}

// getImageFiles walks through the folder and returns sorted image files
func getImageFiles(folderPath string) ([]ImageInfo, error) {
	var imageFiles []string
	err := filepath.Walk(folderPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() {
			fileName := filepath.Base(path)
			imageFiles = append(imageFiles, fileName)
		}
		return nil
	})
	if err != nil {
		return nil, err
	}

	// Sort the image files based on numbers in the filenames
	sort.Slice(imageFiles, func(i, j int) bool {
		re := regexp.MustCompile(`\d+`)
		hasNumI := re.MatchString(imageFiles[i])
		hasNumJ := re.MatchString(imageFiles[j])

		if hasNumI && !hasNumJ {
			return false
		} else if !hasNumI && hasNumJ {
			return true
		}
		return imageFiles[i] < imageFiles[j]
	})

	var images []ImageInfo
	for _, fileName := range imageFiles {
		images = append(images, ImageInfo{FilePath: folderPath + fileName})
	}
	return images, nil
}

// openExcelFile opens the specified Excel template file
func openExcelFile(templatePath string) (*excelize.File, error) {
	return excelize.OpenFile(templatePath)
}

// saveExcelFile saves the Excel file
func saveExcelFile(f *excelize.File) error {
	return f.Save()
}

// pasteImagesHorizontally places images horizontally in the Excel sheet
func pasteImagesHorizontally(f *excelize.File, sheetName string, images []ImageInfo, startCell string) error {
	currentCol, row, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		return fmt.Errorf("invalid starting cell: %v", err)
	}

	desiredWidth := 1115.9 // Desired width in pixels
	desiredHeight := 609.2 // Desired height in pixels

	for index, img := range images {
		cellName, _ := excelize.CoordinatesToCellName(currentCol, row)

		// Get original dimensions of the image
		originalWidth, originalHeight, err := getDimensions(img.FilePath)
		if err != nil {
			return fmt.Errorf("failed to get image dimensions: %v", err)
		}

		// Calculate scaling factors
		scaleX := float64(desiredWidth) / float64(originalWidth)
		scaleY := float64(desiredHeight) / float64(originalHeight)

		// Add the image at the current position
		err = addImage(f, sheetName, img.FilePath, cellName, scaleX, scaleY)
		if err != nil {
			return fmt.Errorf("failed to insert image %s: %v", img.FilePath, err)
		}

		// Move to the next column with spacing
		currentCol += 37

		// Insert a page break after the current image except for the last one
		if index > 0 {
			pageBreakCell, _ := excelize.CoordinatesToCellName(currentCol-1, 40)
			err = f.InsertPageBreak(sheetName, pageBreakCell)
			if err != nil {
				return fmt.Errorf("failed to insert page break at %s: %v", pageBreakCell, err)
			}
		}
	}
	return nil
}

// addImage adds an image at a specific cell in the Excel sheet
func addImage(f *excelize.File, sheetName, filePath, cell string, scaleX, scaleY float64) error {
	imgBytes, err := os.ReadFile(filePath)
	if err != nil {
		return fmt.Errorf("failed to read image file: %v", err)
	}

	err = f.AddPictureFromBytes(sheetName, cell, &excelize.Picture{
		Extension: ".png", // Ensure the file extension matches
		File:      imgBytes,
		Format: &excelize.GraphicOptions{
			ScaleX:  scaleX,
			ScaleY:  scaleY,
			AutoFit: false,
		},
	})
	if err != nil {
		return fmt.Errorf("failed to insert image: %v", err)
	}
	return nil
}

func getDimensions(filePath string) (int, int, error) {
	imgFile, err := os.Open(filePath)
	if err != nil {
		return 0, 0, err
	}
	defer imgFile.Close()
	img, _, err := image.Decode(imgFile)
	if err != nil {
		return 0, 0, err
	}
	return img.Bounds().Max.X, img.Bounds().Max.Y, nil
}
