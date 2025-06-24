package main

import (
	"fmt"
	"math"
	"math/rand"
	"os"
	"sort"
	"time"

	"github.com/kreddevils18/go-excelbuilder/pkg/excelbuilder"
)

type Student struct {
	ID          string
	FirstName   string
	LastName    string
	Email       string
	Grade       int
	DateOfBirth time.Time
	ParentName  string
	ParentEmail string
	ParentPhone string
	Address     string
	EmergencyContact string
	Allergies   string
	Notes       string
}

type Course struct {
	ID          string
	Name        string
	Subject     string
	Teacher     string
	Credits     float64
	Description string
	Room        string
	Period      string
	Semester    string
	Year        int
}

type Assignment struct {
	ID          string
	CourseID    string
	Name        string
	Category    string // Homework, Quiz, Test, Project, Participation
	MaxPoints   float64
	Weight      float64
	DueDate     time.Time
	Description string
	Rubric      string
}

type Grade struct {
	StudentID    string
	AssignmentID string
	PointsEarned float64
	MaxPoints    float64
	Percentage   float64
	LetterGrade  string
	Submitted    time.Time
	Late         bool
	Excused      bool
	Comments     string
}

type Attendance struct {
	StudentID string
	CourseID  string
	Date      time.Time
	Status    string // Present, Absent, Tardy, Excused
	Period    string
	Notes     string
}

type GradeSummary struct {
	StudentID       string
	CourseID        string
	HomeworkAvg     float64
	QuizAvg         float64
	TestAvg         float64
	ProjectAvg      float64
	ParticipationAvg float64
	OverallAvg      float64
	LetterGrade     string
	GPA             float64
	Rank            int
}

func main() {
	fmt.Println("Gradebook Management System Example")
	fmt.Println("==================================")

	// Create output directory
	if err := os.MkdirAll("output", 0755); err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}

	// Generate gradebook reports
	fmt.Println("Generating Student Roster...")
	if err := generateStudentRoster(); err != nil {
		fmt.Printf("Error generating student roster: %v\n", err)
	} else {
		fmt.Println("✓ Student Roster generated")
	}

	fmt.Println("Generating Grade Tracking Report...")
	if err := generateGradeTracking(); err != nil {
		fmt.Printf("Error generating grade tracking: %v\n", err)
	} else {
		fmt.Println("✓ Grade Tracking Report generated")
	}

	fmt.Println("Generating Student Reports...")
	if err := generateStudentReports(); err != nil {
		fmt.Printf("Error generating student reports: %v\n", err)
	} else {
		fmt.Println("✓ Student Reports generated")
	}

	fmt.Println("Generating Analytics Dashboard...")
	if err := generateAnalyticsDashboard(); err != nil {
		fmt.Printf("Error generating analytics: %v\n", err)
	} else {
		fmt.Println("✓ Analytics Dashboard generated")
	}

	fmt.Println("Generating Attendance Report...")
	if err := generateAttendanceReport(); err != nil {
		fmt.Printf("Error generating attendance report: %v\n", err)
	} else {
		fmt.Println("✓ Attendance Report generated")
	}

	fmt.Println("\nGradebook management reports completed!")
	fmt.Println("Check the output directory for generated files.")
}

func generateStudentRoster() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Student Roster")

	// Title and metadata
	sheet.SetCell("A1", "STUDENT ROSTER - FALL 2024").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#2E75B6"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:M1")

	// Class information
	sheet.SetCell("A3", "Class:").SetStyle(getBoldStyle())
	sheet.SetCell("B3", "Mathematics - Algebra II")
	sheet.SetCell("A4", "Teacher:").SetStyle(getBoldStyle())
	sheet.SetCell("B4", "Ms. Sarah Johnson")
	sheet.SetCell("A5", "Room:").SetStyle(getBoldStyle())
	sheet.SetCell("B5", "Room 205")
	sheet.SetCell("A6", "Period:").SetStyle(getBoldStyle())
	sheet.SetCell("B6", "3rd Period")

	// Generate student data
	students := generateStudents()
	createStudentRosterTable(sheet, students, 8)

	// Class statistics
	createClassStatistics(sheet, students, 8+len(students)+3)

	return builder.SaveToFile("output/15-gradebook-roster.xlsx")
}

func generateGradeTracking() error {
	builder := excelbuilder.NewBuilder()

	// Assignments sheet
	assignmentSheet := builder.AddSheet("Assignments")
	createAssignmentsSheet(assignmentSheet)

	// Grades sheet
	gradesSheet := builder.AddSheet("Grades")
	createGradesSheet(gradesSheet)

	// Grade Summary sheet
	summarySheet := builder.AddSheet("Grade Summary")
	createGradeSummarySheet(summarySheet)

	// Gradebook sheet
	gradebookSheet := builder.AddSheet("Gradebook")
	createGradebookSheet(gradebookSheet)

	return builder.SaveToFile("output/15-gradebook-grades.xlsx")
}

func generateStudentReports() error {
	builder := excelbuilder.NewBuilder()

	// Individual Progress Reports
	progressSheet := builder.AddSheet("Progress Reports")
	createProgressReportsSheet(progressSheet)

	// Parent Communication
	parentSheet := builder.AddSheet("Parent Reports")
	createParentReportsSheet(parentSheet)

	// Academic Alerts
	alertsSheet := builder.AddSheet("Academic Alerts")
	createAcademicAlertsSheet(alertsSheet)

	return builder.SaveToFile("output/15-gradebook-reports.xlsx")
}

func generateAnalyticsDashboard() error {
	builder := excelbuilder.NewBuilder()

	// Class Performance
	performanceSheet := builder.AddSheet("Class Performance")
	createClassPerformanceSheet(performanceSheet)

	// Grade Distribution
	distributionSheet := builder.AddSheet("Grade Distribution")
	createGradeDistributionSheet(distributionSheet)

	// Trends Analysis
	trendsSheet := builder.AddSheet("Trends Analysis")
	createTrendsAnalysisSheet(trendsSheet)

	// Standards Mastery
	standardsSheet := builder.AddSheet("Standards Mastery")
	createStandardsMasterySheet(standardsSheet)

	return builder.SaveToFile("output/15-gradebook-analytics.xlsx")
}

func generateAttendanceReport() error {
	builder := excelbuilder.NewBuilder()
	sheet := builder.AddSheet("Attendance")

	// Title
	sheet.SetCell("A1", "ATTENDANCE TRACKING").SetStyle(&excelbuilder.Style{
		Font: &excelbuilder.Font{Bold: true, Size: 16, Color: "#FFFFFF"},
		Fill: &excelbuilder.Fill{Type: "solid", Color: "#70AD47"},
		Alignment: &excelbuilder.Alignment{Horizontal: "center"},
	})
	sheet.MergeRange("A1:H1")

	// Generate attendance data
	attendance := generateAttendanceData()
	createAttendanceTable(sheet, attendance, 3)

	// Attendance summary
	createAttendanceSummary(sheet, attendance, 3+len(attendance)+3)

	return builder.SaveToFile("output/15-gradebook-attendance.xlsx")
}

// Data generation functions
func generateStudents() []Student {
	firstNames := []string{"Emma", "Liam", "Olivia", "Noah", "Ava", "Ethan", "Sophia", "Mason", "Isabella", "William", "Mia", "James", "Charlotte", "Benjamin", "Amelia", "Lucas", "Harper", "Henry", "Evelyn", "Alexander", "Abigail", "Michael", "Emily", "Daniel", "Elizabeth"}
	lastNames := []string{"Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White", "Harris"}

	students := make([]Student, 25)
	for i := 0; i < 25; i++ {
		firstName := firstNames[i]
		lastName := lastNames[i]
		email := fmt.Sprintf("%s.%s@school.edu", strings.ToLower(firstName), strings.ToLower(lastName))
		parentEmail := fmt.Sprintf("%s.parent@email.com", strings.ToLower(lastName))

		students[i] = Student{
			ID:          fmt.Sprintf("STU%03d", i+1),
			FirstName:   firstName,
			LastName:    lastName,
			Email:       email,
			Grade:       10 + rand.Intn(3), // Grades 10-12
			DateOfBirth: time.Date(2006+rand.Intn(3), time.Month(rand.Intn(12)+1), rand.Intn(28)+1, 0, 0, 0, 0, time.UTC),
			ParentName:  fmt.Sprintf("%s %s", []string{"John", "Jane", "Michael", "Sarah", "David", "Lisa"}[rand.Intn(6)], lastName),
			ParentEmail: parentEmail,
			ParentPhone: fmt.Sprintf("(%03d) %03d-%04d", 200+rand.Intn(800), rand.Intn(1000), rand.Intn(10000)),
			Address:     fmt.Sprintf("%d %s Street, City, ST %05d", 100+rand.Intn(9900), []string{"Main", "Oak", "Pine", "Elm", "Maple"}[rand.Intn(5)], 10000+rand.Intn(90000)),
			EmergencyContact: fmt.Sprintf("Emergency Contact: (%03d) %03d-%04d", 200+rand.Intn(800), rand.Intn(1000), rand.Intn(10000)),
			Allergies:   []string{"None", "Peanuts", "Dairy", "Shellfish", "None", "None"}[rand.Intn(6)],
			Notes:       []string{"", "Excellent student", "Needs extra help", "Very creative", "Strong in math", ""}[rand.Intn(6)],
		}
	}

	return students
}

func generateAssignments() []Assignment {
	assignments := []Assignment{
		{"A001", "MATH101", "Homework 1: Linear Equations", "Homework", 20, 0.15, time.Now().AddDate(0, 0, -30), "Solve linear equations", "Standard rubric"},
		{"A002", "MATH101", "Quiz 1: Graphing", "Quiz", 25, 0.20, time.Now().AddDate(0, 0, -25), "Graph linear functions", "Quiz rubric"},
		{"A003", "MATH101", "Test 1: Chapter 1-2", "Test", 100, 0.30, time.Now().AddDate(0, 0, -20), "Comprehensive test", "Test rubric"},
		{"A004", "MATH101", "Project: Real World Applications", "Project", 50, 0.25, time.Now().AddDate(0, 0, -15), "Apply math to real scenarios", "Project rubric"},
		{"A005", "MATH101", "Participation Week 1", "Participation", 10, 0.10, time.Now().AddDate(0, 0, -28), "Class participation", "Participation rubric"},
	}

	// Generate more assignments
	for i := 6; i <= 20; i++ {
		categories := []string{"Homework", "Quiz", "Test", "Project", "Participation"}
		weights := []float64{0.15, 0.20, 0.30, 0.25, 0.10}
		maxPoints := []float64{20, 25, 100, 50, 10}

		categoryIndex := rand.Intn(len(categories))
		category := categories[categoryIndex]

		assignment := Assignment{
			ID:          fmt.Sprintf("A%03d", i),
			CourseID:    "MATH101",
			Name:        fmt.Sprintf("%s %d", category, i),
			Category:    category,
			MaxPoints:   maxPoints[categoryIndex],
			Weight:      weights[categoryIndex],
			DueDate:     time.Now().AddDate(0, 0, -rand.Intn(30)),
			Description: fmt.Sprintf("Description for %s %d", category, i),
			Rubric:      fmt.Sprintf("%s rubric", category),
		}
		assignments = append(assignments, assignment)
	}

	return assignments
}

func generateGrades() []Grade {
	grades := make([]Grade, 0)
	students := generateStudents()
	assignments := generateAssignments()

	for _, student := range students {
		for _, assignment := range assignments {
			// Some students might not have all grades
			if rand.Float64() < 0.95 { // 95% completion rate
				// Generate grade based on student performance level
				performanceLevel := rand.Float64()
				var percentage float64

				switch {
				case performanceLevel < 0.1: // 10% struggling students
					percentage = 50 + rand.Float64()*20 // 50-70%
				case performanceLevel < 0.3: // 20% below average
					percentage = 70 + rand.Float64()*10 // 70-80%
				case performanceLevel < 0.7: // 40% average
					percentage = 80 + rand.Float64()*10 // 80-90%
				default: // 30% above average
					percentage = 90 + rand.Float64()*10 // 90-100%
				}

				pointsEarned := (percentage / 100.0) * assignment.MaxPoints
				letterGrade := calculateLetterGrade(percentage)
				late := rand.Float64() < 0.1 // 10% late submissions

				grade := Grade{
					StudentID:    student.ID,
					AssignmentID: assignment.ID,
					PointsEarned: pointsEarned,
					MaxPoints:    assignment.MaxPoints,
					Percentage:   percentage,
					LetterGrade:  letterGrade,
					Submitted:    assignment.DueDate.Add(time.Duration(rand.Intn(48)) * time.Hour),
					Late:         late,
					Excused:      false,
					Comments:     generateGradeComment(percentage),
				}
				grades = append(grades, grade)
			}
		}
	}

	return grades
}

func generateAttendanceData() []Attendance {
	attendance := make([]Attendance, 0)
	students := generateStudents()

	// Generate 30 days of attendance
	for i := 0; i < 30; i++ {
		date := time.Now().AddDate(0, 0, -i)
		// Skip weekends
		if date.Weekday() == time.Saturday || date.Weekday() == time.Sunday {
			continue
		}

		for _, student := range students {
			// 95% attendance rate
			var status string
			random := rand.Float64()
			switch {
			case random < 0.90:
				status = "Present"
			case random < 0.95:
				status = "Tardy"
			case random < 0.98:
				status = "Excused"
			default:
				status = "Absent"
			}

			attendanceRecord := Attendance{
				StudentID: student.ID,
				CourseID:  "MATH101",
				Date:      date,
				Status:    status,
				Period:    "3rd Period",
				Notes:     generateAttendanceNote(status),
			}
			attendance = append(attendance, attendanceRecord)
		}
	}

	return attendance
}

// Table creation functions
func createStudentRosterTable(sheet *excelbuilder.Sheet, students []Student, startRow int) {
	headers := []string{"Student ID", "First Name", "Last Name", "Grade", "Email", "Date of Birth", "Parent Name", "Parent Email", "Parent Phone", "Address", "Emergency Contact", "Allergies", "Notes"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, student := range students {
		row := startRow + i + 1
		sheet.SetCell(fmt.Sprintf("A%d", row), student.ID)
		sheet.SetCell(fmt.Sprintf("B%d", row), student.FirstName)
		sheet.SetCell(fmt.Sprintf("C%d", row), student.LastName)
		sheet.SetCell(fmt.Sprintf("D%d", row), student.Grade)
		sheet.SetCell(fmt.Sprintf("E%d", row), student.Email)
		sheet.SetCell(fmt.Sprintf("F%d", row), student.DateOfBirth.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("G%d", row), student.ParentName)
		sheet.SetCell(fmt.Sprintf("H%d", row), student.ParentEmail)
		sheet.SetCell(fmt.Sprintf("I%d", row), student.ParentPhone)
		sheet.SetCell(fmt.Sprintf("J%d", row), student.Address)
		sheet.SetCell(fmt.Sprintf("K%d", row), student.EmergencyContact)
		sheet.SetCell(fmt.Sprintf("L%d", row), student.Allergies)
		sheet.SetCell(fmt.Sprintf("M%d", row), student.Notes)

		// Highlight students with allergies
		if student.Allergies != "None" && student.Allergies != "" {
			sheet.SetCell(fmt.Sprintf("L%d", row), student.Allergies).SetStyle(&excelbuilder.Style{
				Fill: &excelbuilder.Fill{Type: "solid", Color: "#FFE6E6"},
				Font: &excelbuilder.Font{Bold: true},
			})
		}
	}
}

func createAttendanceTable(sheet *excelbuilder.Sheet, attendance []Attendance, startRow int) {
	headers := []string{"Student ID", "Date", "Status", "Period", "Notes"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, startRow), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, record := range attendance {
		row := startRow + i + 1
		sheet.SetCell(fmt.Sprintf("A%d", row), record.StudentID)
		sheet.SetCell(fmt.Sprintf("B%d", row), record.Date.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("D%d", row), record.Period)
		sheet.SetCell(fmt.Sprintf("E%d", row), record.Notes)

		// Status with color coding
		statusStyle := getDataCellStyle()
		switch record.Status {
		case "Present":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		case "Tardy":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		case "Absent":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		case "Excused":
			statusStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#E2E3E5"}
		}
		sheet.SetCell(fmt.Sprintf("C%d", row), record.Status).SetStyle(statusStyle)
	}
}

// Helper functions
func calculateLetterGrade(percentage float64) string {
	switch {
	case percentage >= 97:
		return "A+"
	case percentage >= 93:
		return "A"
	case percentage >= 90:
		return "A-"
	case percentage >= 87:
		return "B+"
	case percentage >= 83:
		return "B"
	case percentage >= 80:
		return "B-"
	case percentage >= 77:
		return "C+"
	case percentage >= 73:
		return "C"
	case percentage >= 70:
		return "C-"
	case percentage >= 67:
		return "D+"
	case percentage >= 65:
		return "D"
	default:
		return "F"
	}
}

func generateGradeComment(percentage float64) string {
	switch {
	case percentage >= 95:
		return "Excellent work!"
	case percentage >= 85:
		return "Good job!"
	case percentage >= 75:
		return "Satisfactory"
	case percentage >= 65:
		return "Needs improvement"
	default:
		return "Please see me for help"
	}
}

func generateAttendanceNote(status string) string {
	switch status {
	case "Tardy":
		return "Late arrival"
	case "Absent":
		return "Unexcused absence"
	case "Excused":
		return "Doctor appointment"
	default:
		return ""
	}
}

// Additional sheet creation functions
func createAssignmentsSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "ASSIGNMENTS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:I1")

	assignments := generateAssignments()
	headers := []string{"Assignment ID", "Name", "Category", "Max Points", "Weight", "Due Date", "Description", "Rubric"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, 3), header).SetStyle(getTableHeaderStyle())
	}

	// Data
	for i, assignment := range assignments {
		row := i + 4
		sheet.SetCell(fmt.Sprintf("A%d", row), assignment.ID)
		sheet.SetCell(fmt.Sprintf("B%d", row), assignment.Name)
		sheet.SetCell(fmt.Sprintf("C%d", row), assignment.Category)
		sheet.SetCell(fmt.Sprintf("D%d", row), assignment.MaxPoints)
		sheet.SetCell(fmt.Sprintf("E%d", row), assignment.Weight).SetStyle(&excelbuilder.Style{NumberFormat: "0.0%"})
		sheet.SetCell(fmt.Sprintf("F%d", row), assignment.DueDate.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("G%d", row), assignment.Description)
		sheet.SetCell(fmt.Sprintf("H%d", row), assignment.Rubric)
	}
}

func createGradesSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "INDIVIDUAL GRADES").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:J1")

	grades := generateGrades()
	headers := []string{"Student ID", "Assignment ID", "Points Earned", "Max Points", "Percentage", "Letter Grade", "Submitted", "Late", "Comments"}

	// Headers
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, 3), header).SetStyle(getTableHeaderStyle())
	}

	// Data (show first 100 grades to avoid overwhelming)
	maxGrades := 100
	if len(grades) < maxGrades {
		maxGrades = len(grades)
	}

	for i := 0; i < maxGrades; i++ {
		grade := grades[i]
		row := i + 4
		sheet.SetCell(fmt.Sprintf("A%d", row), grade.StudentID)
		sheet.SetCell(fmt.Sprintf("B%d", row), grade.AssignmentID)
		sheet.SetCell(fmt.Sprintf("C%d", row), grade.PointsEarned).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("D%d", row), grade.MaxPoints).SetStyle(&excelbuilder.Style{NumberFormat: "0.0"})
		sheet.SetCell(fmt.Sprintf("E%d", row), grade.Percentage).SetStyle(&excelbuilder.Style{NumberFormat: "0.0%"})
		sheet.SetCell(fmt.Sprintf("G%d", row), grade.Submitted.Format("2006-01-02"))
		sheet.SetCell(fmt.Sprintf("H%d", row), grade.Late)
		sheet.SetCell(fmt.Sprintf("I%d", row), grade.Comments)

		// Letter grade with color coding
		gradeStyle := getDataCellStyle()
		switch grade.LetterGrade[0] {
		case 'A':
			gradeStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
		case 'B':
			gradeStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#E2F3FF"}
		case 'C':
			gradeStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
		case 'D':
			gradeStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFE6CC"}
		case 'F':
			gradeStyle.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
		}
		sheet.SetCell(fmt.Sprintf("F%d", row), grade.LetterGrade).SetStyle(gradeStyle)
	}
}

// Summary and statistics functions
func createClassStatistics(sheet *excelbuilder.Sheet, students []Student, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "CLASS STATISTICS").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:D%d", startRow, startRow))

	// Calculate statistics
	gradeCount := make(map[int]int)
	for _, student := range students {
		gradeCount[student.Grade]++
	}

	statsData := [][]interface{}{
		{"Metric", "Value", "Details", "Status"},
		{"Total Students", len(students), "Enrolled", "Active"},
		{"Grade 10", gradeCount[10], "Sophomores", "Current"},
		{"Grade 11", gradeCount[11], "Juniors", "Current"},
		{"Grade 12", gradeCount[12], "Seniors", "Current"},
		{"Average Age", "16.2 years", "Calculated", "Current"},
	}

	for i, row := range statsData {
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

func createAttendanceSummary(sheet *excelbuilder.Sheet, attendance []Attendance, startRow int) {
	sheet.SetCell(fmt.Sprintf("A%d", startRow), "ATTENDANCE SUMMARY").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange(fmt.Sprintf("A%d:D%d", startRow, startRow))

	// Calculate attendance statistics
	statusCount := make(map[string]int)
	for _, record := range attendance {
		statusCount[record.Status]++
	}

	total := len(attendance)
	summaryData := [][]interface{}{
		{"Status", "Count", "Percentage", "Notes"},
		{"Present", statusCount["Present"], fmt.Sprintf("%.1f%%", float64(statusCount["Present"])/float64(total)*100), "On time"},
		{"Tardy", statusCount["Tardy"], fmt.Sprintf("%.1f%%", float64(statusCount["Tardy"])/float64(total)*100), "Late arrival"},
		{"Absent", statusCount["Absent"], fmt.Sprintf("%.1f%%", float64(statusCount["Absent"])/float64(total)*100), "Unexcused"},
		{"Excused", statusCount["Excused"], fmt.Sprintf("%.1f%%", float64(statusCount["Excused"])/float64(total)*100), "Excused absence"},
	}

	for i, row := range summaryData {
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

// Additional placeholder functions for complex sheets
func createGradeSummarySheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "GRADE SUMMARY BY STUDENT").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:H1")

	// This would contain calculated averages by category for each student
	headers := []string{"Student ID", "Homework Avg", "Quiz Avg", "Test Avg", "Project Avg", "Participation", "Overall Avg", "Letter Grade"}
	for i, header := range headers {
		col := string(rune('A' + i))
		sheet.SetCell(fmt.Sprintf("%s%d", col, 3), header).SetStyle(getTableHeaderStyle())
	}

	// Sample data
	sampleData := [][]interface{}{
		{"STU001", 92.5, 88.0, 85.5, 94.0, 95.0, 89.2, "B+"},
		{"STU002", 78.5, 82.0, 79.5, 88.0, 85.0, 81.4, "B-"},
		{"STU003", 95.5, 94.0, 92.5, 97.0, 98.0, 94.8, "A"},
	}

	for i, row := range sampleData {
		rowNum := i + 4
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if j >= 1 && j <= 6 {
				style.NumberFormat = "0.0"
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

func createGradebookSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "GRADEBOOK MATRIX").SetStyle(getSectionHeaderStyle())
	sheet.MergeRange("A1:F1")

	// This would be a matrix view with students as rows and assignments as columns
	sheet.SetCell("A3", "Student").SetStyle(getTableHeaderStyle())
	sheet.SetCell("B3", "HW1").SetStyle(getTableHeaderStyle())
	sheet.SetCell("C3", "Quiz1").SetStyle(getTableHeaderStyle())
	sheet.SetCell("D3", "Test1").SetStyle(getTableHeaderStyle())
	sheet.SetCell("E3", "Project1").SetStyle(getTableHeaderStyle())
	sheet.SetCell("F3", "Average").SetStyle(getTableHeaderStyle())

	// Sample gradebook data
	gradebookData := [][]interface{}{
		{"Emma Smith", 18, 23, 85, 47, 89.2},
		{"Liam Johnson", 16, 20, 79, 42, 81.4},
		{"Olivia Williams", 19, 24, 92, 49, 94.8},
	}

	for i, row := range gradebookData {
		rowNum := i + 4
		for j, cell := range row {
			col := string(rune('A' + j))
			style := getDataCellStyle()
			if j >= 1 && j <= 4 {
				style.NumberFormat = "0.0"
			} else if j == 5 {
				style.NumberFormat = "0.0"
				if cell.(float64) >= 90 {
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#D4EDDA"}
				} else if cell.(float64) >= 80 {
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#FFF3CD"}
				} else {
					style.Fill = &excelbuilder.Fill{Type: "solid", Color: "#F8D7DA"}
				}
			}
			sheet.SetCell(fmt.Sprintf("%s%d", col, rowNum), cell).SetStyle(style)
		}
	}
}

// Placeholder functions for additional complex sheets
func createProgressReportsSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "STUDENT PROGRESS REPORTS").SetStyle(getSectionHeaderStyle())
	sheet.SetCell("A3", "Individual progress tracking and trend analysis would be displayed here.")
}

func createParentReportsSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "PARENT COMMUNICATION REPORTS").SetStyle(getSectionHeaderStyle())
	sheet.SetCell("A3", "Parent-friendly reports and communication logs would be displayed here.")
}

func createAcademicAlertsSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "ACADEMIC ALERTS").SetStyle(getSectionHeaderStyle())
	sheet.SetCell("A3", "Students requiring intervention and support would be listed here.")
}

func createClassPerformanceSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "CLASS PERFORMANCE ANALYTICS").SetStyle(getSectionHeaderStyle())
	sheet.SetCell("A3", "Overall class performance metrics and comparisons would be displayed here.")
}

func createGradeDistributionSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "GRADE DISTRIBUTION ANALYSIS").SetStyle(getSectionHeaderStyle())
	sheet.SetCell("A3", "Statistical distribution of grades and performance patterns would be shown here.")
}

func createTrendsAnalysisSheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "PERFORMANCE TRENDS").SetStyle(getSectionHeaderStyle())
	sheet.SetCell("A3", "Longitudinal analysis of student and class performance trends would be displayed here.")
}

func createStandardsMasterySheet(sheet *excelbuilder.Sheet) {
	sheet.SetCell("A1", "STANDARDS MASTERY TRACKING").SetStyle(getSectionHeaderStyle())
	sheet.SetCell("A3", "Alignment with educational standards and mastery levels would be tracked here.")
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