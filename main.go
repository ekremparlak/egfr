package main

import (
	"fmt"
	"math"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	// vars for input
	var kreatinimikromolliterInput string
	var ageInput string
	var genderInput string
	var egfrInput string
	var sheetNameInput string
	var inputFileName string
	var outputFileName string

	fmt.Print("Enter input file name (Default input.xlsx): ")
	fmt.Scanln(&inputFileName)
	if inputFileName == "" {
		inputFileName = "input.xlsx"
	}

	// open file
	f, err := excelize.OpenFile(inputFileName)
	if err != nil {
		panic(err)
	}
	// close file when code ends
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	fmt.Print("Enter sheet name (if empty first sheet will be used): ")
	fmt.Scanln(&sheetNameInput)
	if sheetNameInput == "" {
		sheetNameInput = f.GetSheetName(0)
	}

	fmt.Print("Enter output file name (Default output.xlsx): ")
	fmt.Scanln(&outputFileName)
	if outputFileName == "" {
		outputFileName = "output.xlsx"
	}

	fmt.Print("Enter kreatinimikromolliter column name: ")
	fmt.Scanln(&kreatinimikromolliterInput)
	for kreatinimikromolliterInput == "" {
		fmt.Print("Enter kreatinimikromolliter column name: ")
		fmt.Scanln(&kreatinimikromolliterInput)
	}

	fmt.Print("Enter age column name: ")
	fmt.Scanln(&ageInput)
	for ageInput == "" {
		fmt.Print("Enter age column name: ")
		fmt.Scanln(&ageInput)
	}

	fmt.Print("Enter gender column name: ")
	fmt.Scanln(&genderInput)
	for genderInput == "" {
		fmt.Print("Enter gender column name: ")
		fmt.Scanln(&genderInput)
	}

	fmt.Print("Enter egfr column name: ")
	fmt.Scanln(&egfrInput)
	for egfrInput == "" {
		fmt.Print("Enter egfr column name: ")
		fmt.Scanln(&egfrInput)
	}

	rows, err := f.GetRows(sheetNameInput)
	if err != nil {
		fmt.Println(err)
		return
	}

	for i := range rows {
		if i == 0 || i == 1 {
			continue
		}

		serumStr, err := f.GetCellValue(sheetNameInput, kreatinimikromolliterInput+strconv.Itoa(i))
		if err != nil {
			fmt.Println("Error while getting value from "+kreatinimikromolliterInput+strconv.Itoa(i)+" Error:", err)
			continue
		}
		serum, err := strconv.ParseFloat(serumStr, 64)
		if err != nil {
			fmt.Println("Error while parsing value from "+kreatinimikromolliterInput+strconv.Itoa(i)+" Error:", err)
			continue
		}

		ageStr, err := f.GetCellValue(sheetNameInput, ageInput+strconv.Itoa(i))
		if err != nil {
			fmt.Println("Error while getting value from "+ageInput+ageInput+strconv.Itoa(i)+" Error:", err)
			continue
		}
		age, err := strconv.ParseFloat(ageStr, 64)
		if err != nil {
			fmt.Println("Error while parsing value from "+ageInput+strconv.Itoa(i)+" Error:", err)
			continue
		}

		genderStr, err := f.GetCellValue(sheetNameInput, genderInput+strconv.Itoa(i))
		if err != nil {
			fmt.Println("Error while getting value from "+genderInput+strconv.Itoa(i)+" Error:", err)
			continue
		}
		gender, err := strconv.Atoi(genderStr)
		if err != nil {
			fmt.Println("Error while parsing value from "+genderInput+strconv.Itoa(i)+" Error:", err)
			continue
		}

		response := mdrd(serum, age, gender)

		f.SetCellValue(sheetNameInput, egfrInput+strconv.Itoa(i), response)
	}
	f.SaveAs(outputFileName)
}

func mdrd(serum float64, age float64, gender int) float64 {
	const creatinine = 175.0
	const mgdltoµmol = 88.4

	serumA := serum / mgdltoµmol
	serumPow := math.Pow(serumA, -1.154)
	agePow := math.Pow(age, -0.203)
	genderMultiplier := 1.0
	if gender == 0 {
		genderMultiplier = 0.742
	}

	egfr := creatinine * serumPow * agePow * genderMultiplier
	return egfr
}
