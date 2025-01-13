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
	"strconv"

	"github.com/xuri/excelize/v2"
)

// ImageInfo holds image file details
type ImageInfo struct {
	FilePath string
}

// Processor defines the interface for processing tasks
type Processor interface {
	Process() error
}

// ExcelProcessor processes Excel related tasks
type ExcelProcessor struct {
	f               *excelize.File
	parentFolder    string
	templateSheet   string
	sampleSheetName string
}

// ImageProcessor processes image-related tasks
type ImageProcessor struct {
	f          *excelize.File
	sheetName  string
	folderPath string
}

// AppContext holds the context for the overall application, including the processors
type AppContext struct {
	ExcelProcessor *ExcelProcessor
	ImageProcessor *ImageProcessor
}

// NewExcelProcessor creates a new instance of ExcelProcessor
func NewExcelProcessor(f *excelize.File, parentFolder, templateSheet, sampleSheetName string) *ExcelProcessor {
	return &ExcelProcessor{
		f:               f,
		parentFolder:    parentFolder,
		templateSheet:   templateSheet,
		sampleSheetName: sampleSheetName,
	}
}

// NewImageProcessor creates a new instance of ImageProcessor
func NewImageProcessor(f *excelize.File, sheetName, folderPath string) *ImageProcessor {
	return &ImageProcessor{
		f:          f,
		sheetName:  sheetName,
		folderPath: folderPath,
	}
}

// Process processes Excel related tasks (creating sheets, etc.)
func (ep *ExcelProcessor) Process() error {
	// Get sorted list of child folders
	subFolders, err := GetSubFolders(ep.parentFolder)
	if err != nil {
		return fmt.Errorf("error scanning parent folder: %v", err)
	}

	// Sort folder names numerically
	SortNumeric(subFolders)

	// Walk through the parent folder and process each sorted subfolder
	for _, folder := range subFolders {
		folderPath := filepath.Join(ep.parentFolder, folder)
		ip := NewImageProcessor(ep.f, "#"+folder, folderPath)
		if err := ip.Process(); err != nil {
			return fmt.Errorf("error processing images in folder %s: %v", folder, err)
		}
	}

	return nil
}

// Process processes image-related tasks (inserting images)
func (ip *ImageProcessor) Process() error {
	fmt.Println("Started adding image for folder:", ip.folderPath)
	// Find the sample sheet
	sampleSheetIndex, err := ip.f.GetSheetIndex("Final Template")
	if err != nil || sampleSheetIndex == -1 {
		return fmt.Errorf("failed to find sheet 'Final Template': %v", err)
	}

	// Create a new sheet based on the "Final Template" sheet
	newSheetIndex, err := ip.f.NewSheet("#" + filepath.Base(ip.folderPath))
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %v", err)
	}
	// Copy the content of "Final Template" sheet to the new sheet
	err = ip.f.CopySheet(sampleSheetIndex, newSheetIndex)
	if err != nil {
		return fmt.Errorf("failed to copy 'Final Template' sheet: %v", err)
	}

	// Get sorted image files from the child folder
	imageFiles, err := GetImageFiles(ip.folderPath)
	if err != nil {
		return fmt.Errorf("error getting image files from folder %s: %v", ip.folderPath, err)
	}

	// Start inserting images at a specific row and column
	startCell := "B4"
	err = PasteImagesHorizontally(ip.f, "#"+filepath.Base(ip.folderPath), imageFiles, startCell)
	if err != nil {
		return fmt.Errorf("error inserting images for sheet %s: %v", "#"+filepath.Base(ip.folderPath), err)
	}

	return nil
}

// GetSheetList lists all sheet names in the Excel file and prints them
func GetSheetList(f *excelize.File) ([]string, error) {
	sheets := f.GetSheetList()

	// Print the sheet names to the console
	fmt.Println("\nSheets in the Excel file :", sheets)

	return sheets, nil
}

// GetImageFiles walks through the folder and returns sorted image files
func GetImageFiles(folderPath string) ([]ImageInfo, error) {
	var imageFiles []string
	err := filepath.Walk(folderPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return fmt.Errorf("error accessing file: %v", err)
		}
		if !info.IsDir() {
			imageFiles = append(imageFiles, path)
		}
		return nil
	})
	if err != nil {
		return nil, fmt.Errorf("error walking through folder %s: %v", folderPath, err)
	}

	// Sort the image files based on the desired logic
	sort.Slice(imageFiles, func(i, j int) bool {
		re := regexp.MustCompile(`\d+`)
		hasNumI := re.MatchString(filepath.Base(imageFiles[i]))
		hasNumJ := re.MatchString(filepath.Base(imageFiles[j]))

		// Prioritize files without numbers
		if !hasNumI && hasNumJ {
			return true
		}
		if hasNumI && !hasNumJ {
			return false
		}

		// If both have numbers or neither have numbers, sort lexicographically
		return imageFiles[i] < imageFiles[j]
	})

	var images []ImageInfo
	for _, fileName := range imageFiles {
		images = append(images, ImageInfo{FilePath: fileName})
	}
	return images, nil
}

// PasteImagesHorizontally places images horizontally in the Excel sheet
func PasteImagesHorizontally(f *excelize.File, sheetName string, images []ImageInfo, startCell string) error {
	currentCol, row, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		return fmt.Errorf("invalid starting cell: %v", err)
	}

	desiredWidth := 1115.9 // Desired width in pixels
	desiredHeight := 609.2 // Desired height in pixels

	for index, img := range images {
		cellName, _ := excelize.CoordinatesToCellName(currentCol, row)

		// Get original dimensions of the image
		originalWidth, originalHeight, err := GetDimensions(img.FilePath)
		if err != nil {
			return fmt.Errorf("failed to get image dimensions for %s: %v", img.FilePath, err)
		}

		// Calculate scaling factors
		scaleX := float64(desiredWidth) / float64(originalWidth)
		scaleY := float64(desiredHeight) / float64(originalHeight)

		// Add the image at the current position
		err = InsertImage(f, sheetName, img.FilePath, cellName, scaleX, scaleY)
		if err != nil {
			return fmt.Errorf("failed to insert image %s: %v", img.FilePath, cellName)
		}

		// Move to the next column with spacing
		currentCol += 37

		// Insert a page break after every 37th column
		PageBreakCell, _ := excelize.CoordinatesToCellName(currentCol-1, 40)
		err = AddPageBreak(f, sheetName, PageBreakCell)
		if err != nil {
			return fmt.Errorf("failed to insert page break at %s: %v", PageBreakCell, err)
		}

		fmt.Println("-", "Page break inserted at Column:", PageBreakCell)
		fmt.Println("-", "Image", index+1, "inserted at Cell:", cellName)
		fmt.Print("\n")

	}
	return nil
}

// InsertImage adds an image at a specific cell in the Excel sheet
func InsertImage(f *excelize.File, sheetName, filePath, cell string, scaleX, scaleY float64) error {
	imgBytes, err := os.ReadFile(filePath)
	if err != nil {
		return fmt.Errorf("failed to read image file %s: %v", filePath, err)
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
		return fmt.Errorf("failed to insert image %s at %s: %v", filePath, cell, err)
	}
	return nil
}

// AddPageBreak inserts a page break in the specified cell
func AddPageBreak(f *excelize.File, sheetName, cell string) error {
	return f.InsertPageBreak(sheetName, cell)
}

// GetDimensions retrieves the dimensions (width, height) of the image
func GetDimensions(filePath string) (int, int, error) {
	imgFile, err := os.Open(filePath)
	if err != nil {
		return 0, 0, fmt.Errorf("failed to open image file %s: %v", filePath, err)
	}
	defer imgFile.Close()
	img, _, err := image.Decode(imgFile)
	if err != nil {
		return 0, 0, fmt.Errorf("failed to decode image %s: %v", filePath, err)
	}
	return img.Bounds().Max.X, img.Bounds().Max.Y, nil
}

// SaveExcel saves the Excel file
func SaveExcel(f *excelize.File) error {
	err := f.Save()
	if err != nil {
		return fmt.Errorf("failed to save Excel file: %v", err)
	}
	return nil
}

// SortNumeric sorts folder names numerically
func SortNumeric(folders []string) {
	sort.Slice(folders, func(i, j int) bool {
		numI, errI := strconv.Atoi(folders[i])
		numJ, errJ := strconv.Atoi(folders[j])

		// Handle conversion errors gracefully
		if errI != nil || errJ != nil {
			// Fallback to string comparison if conversion fails
			return folders[i] < folders[j]
		}
		return numI < numJ
	})
}

// GetSubFolders gets the subfolders in the parent directory
func GetSubFolders(parentFolder string) ([]string, error) {
	var subfolders []string
	err := filepath.Walk(parentFolder, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() && path != parentFolder {
			relativePath, err := filepath.Rel(parentFolder, path)
			if err != nil {
				return err
			}
			subfolders = append(subfolders, relativePath)
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return subfolders, nil
}

// SetPageSettings configures the page layout settings for a sheet.
func SetPageSettings(f *excelize.File, sheetName string) error {
	// Set the orientation to landscape
	orientation := "landscape" // Use string directly for landscape orientation
	err := f.SetPageLayout(sheetName, &excelize.PageLayoutOptions{
		Orientation: &orientation, // Use "landscape" for landscape orientation
		Size:        IntPtr(9),
		FitToHeight: IntPtr(1),
		FitToWidth:  IntPtr(0),
	})
	if err != nil {
		return fmt.Errorf("failed to set page layout: %v", err)
	}

	return nil
}

// IntPtr is a helper function to create a pointer to an int
func IntPtr(i int) *int {
	return &i
}

// After processing the sheets and images, rename sheet "#0" to "#Preparation"
func RenamePreparationSheet(f *excelize.File) error {
	// Check if sheet #0 exists
	sheetIndex, err := f.GetSheetIndex("#0")
	if err != nil || sheetIndex == -1 {
		return fmt.Errorf("sheet #0 not found")
	}

	// Rename sheet #0 to #Preparation
	err = f.SetSheetName("#0", "#Preparation")
	if err != nil {
		return fmt.Errorf("failed to rename sheet #0 to #Preparation: %v", err)
	}

	return nil
}
func main() {
	// Define flags for the parent folder path and Excel template path
	parentFolderPath := flag.String("f", "", "Path to the parent folder containing child folders")
	templatePath := flag.String("e", "", "Path to the Excel template")

	// Parse the command-line flags
	flag.Parse()

	// Validate inputs
	if *parentFolderPath == "" || *templatePath == "" {
		fmt.Println("Please provide the parent folder path using the '-f' flag")
		fmt.Println("Please provide the Excel template path using the '-e' flag.")
		return
	}

	// Open the Excel file
	f, err := excelize.OpenFile(*templatePath)
	if err != nil {
		fmt.Printf("Failed to open template file: %v\n", err)
		return
	}

	// Create an ExcelProcessor for processing Excel tasks
	ep := NewExcelProcessor(f, *parentFolderPath, *templatePath, "Final Template")

	// Start processing Excel related tasks
	err = ep.Process()
	if err != nil {
		fmt.Printf("Error processing Excel tasks: %v\n", err)
		return
	}

	// Rename sheet "#0" to "#Preparation"
	err = RenamePreparationSheet(f)
	if err != nil {
		fmt.Printf("Error renaming sheet: %v\n", err)
		return
	}

	// Set page settings for each sheet
	sheets, err := GetSheetList(f)
	if err != nil {
		fmt.Printf("Error getting sheet list: %v\n", err)
		return
	}

	// Iterate through sheets and set page settings
	for _, sheetName := range sheets {
		err = SetPageSettings(f, sheetName)
		if err != nil {
			fmt.Printf("Error setting page layout for sheet %s: %v\n", sheetName, err)
			return
		}
	}

	// Save the Excel file
	err = SaveExcel(f)
	if err != nil {
		fmt.Printf("Failed to save updated file: %v\n", err)
		return
	}

	fmt.Println("All sheets and images processed successfully.")
}
