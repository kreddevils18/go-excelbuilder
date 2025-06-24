package main

import (
	"fmt"
	"math"
	"math/rand"
	"os"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

type ExperimentalData struct {
	SampleID    string
	Measurement float64
	Uncertainty float64
	Unit        string
	Timestamp   time.Time
	Operator    string
	Instrument  string
	Conditions  string
}

type StatisticalResult struct {
	Parameter   string
	Value       float64
	StdError    float64
	Confidence  float64
	PValue      float64
	Significant bool
	Method      string
}

type ChemicalCompound struct {
	Formula        string
	MolecularWeight float64
	MeltingPoint   float64
	BoilingPoint   float64
	Density        float64
	Solubility     string
	HazardClass    string
}

type SpectroscopicData struct {
	Wavelength  float64
	Absorbance  float64
	Transmittance float64
	Intensity   float64
	PeakType    string
	Assignment  string
}

func main() {
	fmt.Println("Scientific Data Analysis Example")
	fmt.Println("================================")

	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}

	// Generate scientific reports
	fmt.Println("Generating Experimental Data Report...")
	if err := generateExperimentalReport(); err != nil {
		fmt.Printf("Error generating experimental report: %v\n", err)
	} else {
		fmt.Println("✓ Experimental Data Report generated")
	}

	fmt.Println("Generating Statistical Analysis Report...")
	if err := generateStatisticalReport(); err != nil {
		fmt.Printf("Error generating statistical report: %v\n", err)
	} else {
		fmt.Println("✓ Statistical Analysis Report generated")
	}

	fmt.Println("Generating Research Documentation...")
	if err := generateResearchReport(); err != nil {
		fmt.Printf("Error generating research report: %v\n", err)
	} else {
		fmt.Println("✓ Research Documentation generated")
	}

	fmt.Println("Generating Laboratory Results...")
	if err := generateLabResults(); err != nil {
		fmt.Printf("Error generating lab results: %v\n", err)
	} else {
		fmt.Println("✓ Laboratory Results generated")
	}

	fmt.Println("Generating Scientific Calculations...")
	if err := generateCalculationsReport(); err != nil {
		fmt.Printf("Error generating calculations report: %v\n", err)
	} else {
		fmt.Println("✓ Scientific Calculations generated")
	}

	fmt.Println("\nScientific data analysis examples completed!")
	fmt.Println("Check the output directory for generated files.")
}

func generateExperimentalReport() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Experimental Data")

	// Title and metadata
	sheet.SetCell("A1", "EXPERIMENTAL DATA ANALYSIS").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#2E75B6"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H1")

	// Experiment metadata
	sheet.SetCell("A3", "Experiment ID:").SetStyle(getBoldStyle())
	sheet.SetCell("B3", "EXP-2024-001")
	sheet.SetCell("A4", "Date:").SetStyle(getBoldStyle())
	sheet.SetCell("B4", time.Now().Format("2006-01-02"))
	sheet.SetCell("A5", "Principal Investigator:").SetStyle(getBoldStyle())
	sheet.SetCell("B5", "Dr. Jane Smith")
	sheet.SetCell("A6", "Laboratory:").SetStyle(getBoldStyle())
	sheet.SetCell("B6", "Advanced Materials Lab")

	// Data table
	data := generateExperimentalData()
	createExperimentalDataTable(sheet, data, 8)

	// Statistical summary
	createStatisticalSummary(sheet, data, 8+len(data)+3)

	return builder.SaveToFile("output/13-scientific-experiment.xlsx")
}

func generateStatisticalReport() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Statistical Analysis")

	// Title
	sheet.SetCell("A1", "STATISTICAL ANALYSIS REPORT").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#70AD47"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:G1")

	// Statistical results
	results := generateStatisticalResults()
	createStatisticalTable(sheet, results, 3)

	// Hypothesis testing
	createHypothesisTestingSection(sheet, 3+len(results)+3)

	// Correlation analysis
	createCorrelationAnalysis(sheet, 3+len(results)+10)

	return builder.SaveToFile("output/13-scientific-statistics.xlsx")
}

func generateResearchReport() error {
	builder := excelbuilder.NewBuilder()

	// Abstract sheet
	abstractSheet := builder.AddSheet("Abstract")
	createAbstractSheet(abstractSheet)

	// Methodology sheet
	methodSheet := builder.AddSheet("Methodology")
	createMethodologySheet(methodSheet)

	// Results sheet
	resultsSheet := builder.AddSheet("Results")
	createResultsSheet(resultsSheet)

	// Discussion sheet
	discussionSheet := builder.AddSheet("Discussion")
	createDiscussionSheet(discussionSheet)

	// References sheet
	referencesSheet := builder.AddSheet("References")
	createReferencesSheet(referencesSheet)

	return builder.SaveToFile("output/13-scientific-research.xlsx")
}

func generateLabResults() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Laboratory Results")

	// Title
	sheet.SetCell("A1", "LABORATORY ANALYSIS RESULTS").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#C5504B"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:G1")

	// Chemical compounds analysis
	compounds := generateChemicalData()
	createChemicalAnalysisTable(sheet, compounds, 3)

	// Spectroscopic data
	spectroData := generateSpectroscopicData()
	createSpectroscopicTable(sheet, spectroData, 3+len(compounds)+3)

	// Quality control
	createQualityControlSection(sheet, 3+len(compounds)+len(spectroData)+6)

	return builder.SaveToFile("output/13-scientific-lab-results.xlsx")
}

func generateCalculationsReport() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Scientific Calculations")

	// Title
	sheet.SetCell("A1", "SCIENTIFIC CALCULATIONS").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#7030A0"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	// Physical constants
	createPhysicalConstants(sheet, 3)

	// Unit conversions
	createUnitConversions(sheet, 15)

	// Mathematical formulas
	createMathematicalFormulas(sheet, 25)

	// Error propagation
	createErrorPropagation(sheet, 35)

	return builder.SaveToFile("output/13-scientific-calculations.xlsx")
}

// Data generation functions
func generateExperimentalData() []ExperimentalData {
	data := make([]ExperimentalData, 50)
	instruments := []string{"Spectrometer A", "Spectrometer B", "Balance C", "pH Meter D"}
	operators := []string{"Dr. Smith", "Dr. Johnson", "Lab Tech A", "Lab Tech B"}
	conditions := []string{"Room Temp", "Heated", "Cooled", "Vacuum"}

	for i := 0; i < 50; i++ {
		baseValue := 100.0 + rand.NormFloat64()*10
		uncertainty := math.Abs(rand.NormFloat64() * 2.5)

		data[i] = ExperimentalData{
			SampleID:    fmt.Sprintf("S-%03d", i+1),
			Measurement: baseValue,
			Uncertainty: uncertainty,
			Unit:        "mg/L",
			Timestamp:   time.Now().Add(-time.Duration(i) * time.Hour),
			Operator:    operators[rand.Intn(len(operators))],
			Instrument:  instruments[rand.Intn(len(instruments))],
			Conditions:  conditions[rand.Intn(len(conditions))],
		}
	}

	return data
}

func generateStatisticalResults() []StatisticalResult {
	return []StatisticalResult{
		{"Mean", 102.45, 1.23, 95.0, 0.001, true, "t-test"},
		{"Standard Deviation", 8.76, 0.87, 95.0, 0.000, true, "F-test"},
		{"Median", 101.23, 1.45, 95.0, 0.002, true, "Wilcoxon"},
		{"Skewness", 0.12, 0.34, 95.0, 0.723, false, "Jarque-Bera"},
		{"Kurtosis", 2.98, 0.67, 95.0, 0.956, false, "Jarque-Bera"},
		{"Correlation", 0.87, 0.05, 95.0, 0.000, true, "Pearson"},
	}
}

func generateChemicalData() []ChemicalCompound {
	return []ChemicalCompound{
		{"H2SO4", 98.079, -10.0, 337.0, 1.84, "Miscible", "Corrosive"},
		{"NaCl", 58.443, 801.0, 1465.0, 2.16, "360 g/L", "Non-hazardous"},
		{"C6H12O6", 180.156, 146.0, 0.0, 1.54, "909 g/L", "Non-hazardous"},
		{"C2H5OH", 46.068, -114.1, 78.4, 0.789, "Miscible", "Flammable"},
		{"CaCO3", 100.087, 825.0, 0.0, 2.71, "0.013 g/L", "Non-hazardous"},
	}
}

func generateSpectroscopicData() []SpectroscopicData {
	data := make([]SpectroscopicData, 20)
	peakTypes := []string{"Sharp", "Broad", "Medium", "Weak"}
	assignments := []string{"C-H stretch", "O-H stretch", "C=O stretch", "N-H bend", "C-C stretch"}

	for i := 0; i < 20; i++ {
		wavelength := 200.0 + float64(i)*20.0 + rand.Float64()*10.0
		absorbance := rand.Float64() * 2.0
		transmittance := math.Pow(10, -absorbance) * 100
		intensity := absorbance * 1000

		data[i] = SpectroscopicData{
			Wavelength:    wavelength,
			Absorbance:    absorbance,
			Transmittance: transmittance,
			Intensity:     intensity,
			PeakType:      peakTypes[rand.Intn(len(peakTypes))],
			Assignment:    assignments[rand.Intn(len(assignments))],
		}
	}

	return data
}

// Table creation functions
func createExperimentalDataTable(sheet *excelbuilder.Sheet, data []ExperimentalData, startRow int) {
	headers := []string{"Sample ID", "Measurement", "Uncertainty", "Unit", "Timestamp", "Operator", "Instrument", "Conditions"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, item := range data {
		row := startRow + i + 1
		sheet.SetCell(fmt.Sprintf("A%d", row), item.SampleID)
		sheet.SetCell(fmt.Sprintf("B%d", row), item.Measurement).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		sheet.SetCell(fmt.Sprintf("C%d", row), item.Uncertainty).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		sheet.SetCell(fmt.Sprintf("D%d", row), item.Unit)
		sheet.SetCell(fmt.Sprintf("E%d", row), item.Timestamp.Format("2006-01-02 15:04"))
		sheet.SetCell(fmt.Sprintf("F%d", row), item.Operator)
		sheet.SetCell(fmt.Sprintf("G%d", row), item.Instrument)
		sheet.SetCell(fmt.Sprintf("H%d", row), item.Conditions)
	}
}

func createStatisticalSummary(sheet *excelbuilder.Sheet, data []ExperimentalData, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "STATISTICAL SUMMARY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:D%d", startRow, startRow))

	// Calculate statistics
	var sum, sumSq float64
	n := float64(len(data))
	for _, item := range data {
		sum += item.Measurement
		sumSq += item.Measurement * item.Measurement
	}

	mean := sum / n
	variance := (sumSq - sum*sum/n) / (n - 1)
	stdDev := math.Sqrt(variance)
	stdError := stdDev / math.Sqrt(n)

	statsData := [][]interface{}{
		{"Parameter", "Value", "Unit", "Description"},
		{"Sample Size", n, "count", "Number of measurements"},
		{"Mean", mean, "mg/L", "Average value"},
		{"Standard Deviation", stdDev, "mg/L", "Measure of spread"},
		{"Standard Error", stdError, "mg/L", "Error in mean"},
		{"Coefficient of Variation", (stdDev/mean)*100, "%", "Relative variability"},
	}

	for i, row := range statsData {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			cellStyle := getDataCellStyle()
			if i == 0 {
				cellStyle = getTableHeaderStyle()
			} else if j == 1 && i > 0 {
				cellStyle.NumberFormat = "0.000"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(cellStyle)
		}
	}
}

func createStatisticalTable(sheet *excelbuilder.Sheet, results []StatisticalResult, startRow int) {
	headers := []string{"Parameter", "Value", "Std Error", "Confidence %", "P-Value", "Significant", "Method"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, result := range results {
		row := startRow + i + 1
		sheet.SetCell(fmt.Sprintf("A%d", row), result.Parameter)
		sheet.SetCell(fmt.Sprintf("B%d", row), result.Value).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		sheet.SetCell(fmt.Sprintf("C%d", row), result.StdError).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		sheet.SetCell(fmt.Sprintf("D%d", row), result.Confidence).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("E%d", row), result.PValue).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		
		// Significance with conditional formatting
		sigStyle := getDataCellStyle()
		if result.Significant {
			sigStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
			sheet.SetCell(fmt.Sprintf("F%d", row), "Yes").SetStyle(sigStyle)
		} else {
			sigStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
			sheet.SetCell(fmt.Sprintf("F%d", row), "No").SetStyle(sigStyle)
		}
		
		sheet.SetCell(fmt.Sprintf("G%d", row), result.Method)
	}
}

func createHypothesisTestingSection(sheet *excelbuilder.Sheet, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "HYPOTHESIS TESTING").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:E%d", startRow, startRow))

	hypotheses := [][]interface{}{
		{"Test", "Null Hypothesis", "Alternative", "Test Statistic", "Decision"},
		{"Normality", "Data is normally distributed", "Data is not normal", "W = 0.987", "Fail to reject H0"},
		{"Mean = 100", "μ = 100", "μ ≠ 100", "t = 2.45", "Reject H0"},
		{"Variance", "σ² = 64", "σ² ≠ 64", "χ² = 45.2", "Fail to reject H0"},
	}

	for i, row := range hypotheses {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createCorrelationAnalysis(sheet *excelbuilder.Sheet, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "CORRELATION ANALYSIS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:F%d", startRow, startRow))

	correlations := [][]interface{}{
		{"Variable 1", "Variable 2", "Correlation", "P-Value", "Significance", "Interpretation"},
		{"Temperature", "Reaction Rate", 0.87, 0.001, "***", "Strong positive"},
		{"pH", "Yield", -0.65, 0.023, "*", "Moderate negative"},
		{"Pressure", "Selectivity", 0.23, 0.234, "ns", "Weak positive"},
		{"Time", "Conversion", 0.92, 0.000, "***", "Very strong positive"},
	}

	for i, row := range correlations {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j == 2 {
				style.NumberFormat = "0.000"
			} else if j == 3 {
				style.NumberFormat = "0.000"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createChemicalAnalysisTable(sheet *excelbuilder.Sheet, compounds []ChemicalCompound, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "CHEMICAL COMPOUND ANALYSIS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:G%d", startRow, startRow))

	headers := []string{"Formula", "Mol. Weight", "M.P. (°C)", "B.P. (°C)", "Density", "Solubility", "Hazard"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow+2), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, compound := range compounds {
		row := startRow + i + 3
		sheet.SetCell(fmt.Sprintf("A%d", row), compound.Formula)
		sheet.SetCell(fmt.Sprintf("B%d", row), compound.MolecularWeight).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		sheet.SetCell(fmt.Sprintf("C%d", row), compound.MeltingPoint).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("D%d", row), compound.BoilingPoint).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("E%d", row), compound.Density).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		sheet.SetCell(fmt.Sprintf("F%d", row), compound.Solubility)
		
		// Hazard with conditional formatting
		hazardStyle := getDataCellStyle()
		switch compound.HazardClass {
		case "Corrosive":
			hazardStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		case "Flammable":
			hazardStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		case "Non-hazardous":
			hazardStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		}
		sheet.SetCell(fmt.Sprintf("G%d", row), compound.HazardClass).SetStyle(hazardStyle)
	}
}

func createSpectroscopicTable(sheet *excelbuilder.Sheet, data []SpectroscopicData, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "SPECTROSCOPIC DATA").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:F%d", startRow, startRow))

	headers := []string{"Wavelength (nm)", "Absorbance", "Transmittance (%)", "Intensity", "Peak Type", "Assignment"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow+2), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, item := range data {
		row := startRow + i + 3
		sheet.SetCell(fmt.Sprintf("A%d", row), item.Wavelength).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("B%d", row), item.Absorbance).SetStyle(&excelbuilder.Style{NumberFormat: "0.000"})
		sheet.SetCell(fmt.Sprintf("C%d", row), item.Transmittance).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("D%d", row), item.Intensity).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("E%d", row), item.PeakType)
		sheet.SetCell(fmt.Sprintf("F%d", row), item.Assignment)
	}
}

func createQualityControlSection(sheet *excelbuilder.Sheet, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "QUALITY CONTROL").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:E%d", startRow, startRow))

	qcData := [][]interface{}{
		{"Parameter", "Specification", "Measured", "Status", "Notes"},
		{"Calibration Check", "R² > 0.999", "R² = 0.9995", "PASS", "Linear range verified"},
		{"Blank Analysis", "< 0.001 AU", "0.0003 AU", "PASS", "No contamination"},
		{"Duplicate Precision", "RSD < 2%", "RSD = 1.2%", "PASS", "Good reproducibility"},
		{"Spike Recovery", "95-105%", "98.5%", "PASS", "Accurate method"},
		{"Reference Standard", "±5% of certified", "+2.1%", "PASS", "Within tolerance"},
	}

	for i, row := range qcData {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j == 3 && cell == "PASS" {
				style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
			} else if j == 3 && cell == "FAIL" {
				style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

// Research documentation functions
func createAbstractSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "RESEARCH ABSTRACT").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 18, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#1F4E79"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:F1")

	abstractContent := [][]interface{}{
		{"Title:", "Advanced Spectroscopic Analysis of Novel Compounds"},
		{"Authors:", "Dr. Jane Smith, Dr. John Doe, Dr. Mary Johnson"},
		{"Institution:", "University Research Laboratory"},
		{"Date:", time.Now().Format("January 2, 2006")},
		{"", ""},
		{"Objective:", "To investigate the spectroscopic properties of newly synthesized compounds"},
		{"Methods:", "UV-Vis, IR, and NMR spectroscopy were employed for characterization"},
		{"Results:", "All compounds showed characteristic absorption patterns with high purity"},
		{"Conclusion:", "The synthesized compounds exhibit promising properties for further study"},
		{"Keywords:", "spectroscopy, synthesis, characterization, novel compounds"},
	}

	for i, row := range abstractContent {
		rowNum := i + 3
		if len(row) > 1 {
			sheet.SetCell(fmt.Sprintf("A%d", rowNum), row[0]).SetStyle(getBoldStyle())
			sheet.SetCell(fmt.Sprintf("B%d", rowNum), row[1])
			sheet.MergeRange(fmt.Sprintf("B%d:F%d", rowNum, rowNum))
		}
	}
}

func createMethodologySheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "METHODOLOGY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	methodology := [][]interface{}{
		{"1. Sample Preparation", ""},
		{"   - Synthesis conditions:", "Reflux at 80°C for 4 hours"},
		{"   - Purification:", "Column chromatography (silica gel)"},
		{"   - Drying:", "Vacuum oven at 60°C overnight"},
		{"", ""},
		{"2. Instrumentation", ""},
		{"   - UV-Vis:", "Shimadzu UV-2600 (200-800 nm)"},
		{"   - IR:", "Bruker FTIR (4000-400 cm⁻¹)"},
		{"   - NMR:", "Bruker 400 MHz (¹H and ¹³C)"},
		{"", ""},
		{"3. Analysis Conditions", ""},
		{"   - Solvent:", "DMSO-d6 for NMR, CHCl3 for UV-Vis"},
		{"   - Temperature:", "25°C ± 1°C"},
		{"   - Concentration:", "1.0 × 10⁻⁴ M for UV-Vis"},
	}

	for i, row := range methodology {
		rowNum := i + 3
		if len(row) > 1 && row[1] != "" {
			sheet.SetCell(fmt.Sprintf("A%d", rowNum), row[0])
			sheet.SetCell(fmt.Sprintf("B%d", rowNum), row[1])
			sheet.MergeRange(fmt.Sprintf("B%d:F%d", rowNum, rowNum))
		} else {
			sheet.SetCell(fmt.Sprintf("A%d", rowNum), row[0]).SetStyle(getBoldStyle())
			sheet.MergeRange(fmt.Sprintf("A%d:F%d", rowNum, rowNum))
		}
	}
}

func createResultsSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "RESULTS AND DISCUSSION").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	// Key findings
	sheet.SetCell("A3", "KEY FINDINGS").SetStyle(getBoldStyle())
	sheet.MergeRange("A3:F3")

	findings := []string{
		"• All synthesized compounds showed high purity (>95%)",
		"• UV-Vis spectra revealed characteristic π→π* transitions",
		"• IR spectra confirmed expected functional groups",
		"• NMR data consistent with proposed structures",
		"• No significant decomposition observed under analysis conditions",
	}

	for i, finding := range findings {
		sheet.SetCell(fmt.Sprintf("A%d", i+4), finding)
		sheet.MergeRange(fmt.Sprintf("A%d:F%d", i+4, i+4))
	}

	// Statistical summary
	sheet.SetCell("A10", "STATISTICAL SUMMARY").SetStyle(getBoldStyle())
	sheet.MergeRange("A10:F10")

	statsData := [][]interface{}{
		{"Parameter", "Mean", "Std Dev", "Min", "Max", "n"},
		{"Yield (%)", 87.3, 5.2, 78.1, 94.6, 15},
		{"Purity (%)", 96.8, 1.8, 93.2, 99.1, 15},
		{"λmax (nm)", 345.2, 12.4, 328.1, 367.8, 15},
	}

	for i, row := range statsData {
		rowNum := i + 11
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j > 0 && j < 5 {
				style.NumberFormat = "0.0"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createDiscussionSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "DISCUSSION").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	discussion := [][]interface{}{
		{"Significance of Results", ""},
		{"The high yields and purities obtained demonstrate the robustness of the synthetic method.", ""},
		{"The spectroscopic data confirm the successful synthesis of target compounds.", ""},
		{"", ""},
		{"Comparison with Literature", ""},
		{"Our results are consistent with similar studies (Smith et al., 2023).", ""},
		{"The observed λmax values align with theoretical predictions.", ""},
		{"", ""},
		{"Limitations", ""},
		{"Sample size was limited to 15 compounds due to time constraints.", ""},
		{"Long-term stability studies were not conducted.", ""},
		{"", ""},
		{"Future Work", ""},
		{"Expand the study to include more structural variants.", ""},
		{"Investigate biological activity of synthesized compounds.", ""},
		{"Conduct computational studies to support experimental findings.", ""},
	}

	for i, row := range discussion {
		rowNum := i + 3
		if row[0] != "" {
			if row[1] == "" {
				sheet.SetCell(fmt.Sprintf("A%d", rowNum), row[0]).SetStyle(getBoldStyle())
			} else {
				sheet.SetCell(fmt.Sprintf("A%d", rowNum), row[0])
			}
			sheet.MergeRange(fmt.Sprintf("A%d:F%d", rowNum, rowNum))
		}
	}
}

func createReferencesSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "REFERENCES").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	references := []string{
		"1. Smith, J.A., Johnson, M.B., & Brown, C.D. (2023). Advanced spectroscopic techniques in organic synthesis. Journal of Analytical Chemistry, 45(3), 123-135.",
		"2. Wilson, P.K., & Davis, L.M. (2022). Novel approaches to compound characterization. Spectroscopy Today, 18(7), 45-52.",
		"3. Anderson, R.S., Thompson, K.L., & White, S.J. (2023). Statistical analysis in chemical research. Chemical Statistics Quarterly, 12(2), 78-89.",
		"4. Garcia, M.E., & Rodriguez, A.F. (2022). Quality control in analytical laboratories. Laboratory Management, 29(4), 234-241.",
		"5. Lee, H.K., Park, S.Y., & Kim, J.W. (2023). Modern instrumentation for chemical analysis. Instrumental Analysis Review, 31(1), 12-28.",
	}

	for i, ref := range references {
		sheet.SetCell(fmt.Sprintf("A%d", i+3), ref)
		sheet.MergeRange(fmt.Sprintf("A%d:F%d", i+3, i+3))
	}
}

// Scientific calculations functions
func createPhysicalConstants(sheet *excelbuilder.Sheet, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "PHYSICAL CONSTANTS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:F%d", startRow, startRow))

	constants := [][]interface{}{
		{"Constant", "Symbol", "Value", "Unit", "Uncertainty", "Reference"},
		{"Speed of light", "c", 2.99792458e8, "m/s", "exact", "CODATA 2018"},
		{"Planck constant", "h", 6.62607015e-34, "J⋅s", "exact", "CODATA 2018"},
		{"Avogadro constant", "NA", 6.02214076e23, "mol⁻¹", "exact", "CODATA 2018"},
		{"Gas constant", "R", 8.314462618, "J/(mol⋅K)", "exact", "CODATA 2018"},
		{"Boltzmann constant", "kB", 1.380649e-23, "J/K", "exact", "CODATA 2018"},
		{"Elementary charge", "e", 1.602176634e-19, "C", "exact", "CODATA 2018"},
		{"Electron mass", "me", 9.1093837015e-31, "kg", "3.0e-40", "CODATA 2018"},
		{"Proton mass", "mp", 1.67262192369e-27, "kg", "5.1e-37", "CODATA 2018"},
	}

	for i, row := range constants {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			} else if j == 2 {
				style.NumberFormat = "0.00E+00"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createUnitConversions(sheet *excelbuilder.Sheet, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "UNIT CONVERSIONS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:E%d", startRow, startRow))

	conversions := [][]interface{}{
		{"From", "To", "Factor", "Formula", "Example"},
		{"°C", "K", "+273.15", "K = °C + 273.15", "25°C = 298.15 K"},
		{"°F", "°C", "(°F-32)×5/9", "°C = (°F-32)×5/9", "77°F = 25°C"},
		{"atm", "Pa", "×101325", "Pa = atm × 101325", "1 atm = 101325 Pa"},
		{"cal", "J", "×4.184", "J = cal × 4.184", "1 cal = 4.184 J"},
		{"eV", "J", "×1.602e-19", "J = eV × 1.602e-19", "1 eV = 1.602e-19 J"},
		{"Å", "m", "×1e-10", "m = Å × 1e-10", "1 Å = 1e-10 m"},
		{"ppm", "mg/L", "×1", "mg/L = ppm × 1", "10 ppm = 10 mg/L"},
	}

	for i, row := range conversions {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createMathematicalFormulas(sheet *excelbuilder.Sheet, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "MATHEMATICAL FORMULAS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:D%d", startRow, startRow))

	formulas := [][]interface{}{
		{"Application", "Formula", "Variables", "Units"},
		{"Beer-Lambert Law", "A = ε × c × l", "A=absorbance, ε=molar absorptivity, c=concentration, l=path length", "A=dimensionless, ε=L/(mol⋅cm), c=mol/L, l=cm"},
		{"Ideal Gas Law", "PV = nRT", "P=pressure, V=volume, n=moles, R=gas constant, T=temperature", "P=Pa, V=m³, n=mol, R=8.314 J/(mol⋅K), T=K"},
		{"Henderson-Hasselbalch", "pH = pKa + log([A⁻]/[HA])", "pH=acidity, pKa=acid dissociation constant, [A⁻]=conjugate base, [HA]=acid", "pH=dimensionless, concentrations in mol/L"},
		{"Arrhenius Equation", "k = A × exp(-Ea/RT)", "k=rate constant, A=pre-exponential factor, Ea=activation energy", "k=varies, A=varies, Ea=J/mol, R=8.314 J/(mol⋅K), T=K"},
		{"Nernst Equation", "E = E° - (RT/nF)ln(Q)", "E=cell potential, E°=standard potential, n=electrons, F=Faraday constant, Q=reaction quotient", "E=V, R=8.314 J/(mol⋅K), T=K, n=dimensionless, F=96485 C/mol"},
	}

	for i, row := range formulas {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}

	// Set column widths for better readability
	sheet.SetColumnWidth("A", 20)
	sheet.SetColumnWidth("B", 25)
	sheet.SetColumnWidth("C", 40)
	sheet.SetColumnWidth("D", 40)
}

func createErrorPropagation(sheet *excelbuilder.Sheet, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "ERROR PROPAGATION").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:E%d", startRow, startRow))

	errorData := [][]interface{}{
		{"Operation", "Formula", "Error Formula", "Example", "Result"},
		{"Addition/Subtraction", "z = x ± y", "σz = √(σx² + σy²)", "(10.0±0.1) + (5.0±0.2)", "15.0±0.22"},
		{"Multiplication/Division", "z = x × y or z = x / y", "σz/z = √((σx/x)² + (σy/y)²)", "(10.0±0.1) × (5.0±0.2)", "50.0±2.2"},
		{"Power", "z = x^n", "σz/z = n × (σx/x)", "(10.0±0.1)²", "100.0±2.0"},
		{"Logarithm", "z = ln(x)", "σz = σx/x", "ln(10.0±0.1)", "2.303±0.010"},
		{"Exponential", "z = e^x", "σz = z × σx", "e^(2.0±0.1)", "7.39±0.74"},
	}

	for i, row := range errorData {
		rowNum := startRow + i + 2
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if i == 0 {
				style = getTableHeaderStyle()
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}

	// Set column widths
	sheet.SetColumnWidth("A", 18)
	sheet.SetColumnWidth("B", 20)
	sheet.SetColumnWidth("C", 25)
	sheet.SetColumnWidth("D", 20)
	sheet.SetColumnWidth("E", 15)
}

// Style helper functions
func getBoldStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true},
	}
}

func getSectionHeaderStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 14, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#2E75B6"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center", Vertical: "center"},
	}
}

func getTableHeaderStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#4472C4"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center", Vertical: "center"},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
			Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#000000"},
		},
	}
}

func getDataCellStyle() *excelbuilder.Style {
	return &excelbuilder.Style{
		Alignment: &excelbuilder.Alignment{Horizontal: "left", Vertical: "center"},
		Border: &excelbuilder.Border{
			Top:    &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Bottom: &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Left:   &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
			Right:  &excelbuilder.BorderStyle{Style: "thin", Color: "#CCCCCC"},
		},
	}
}