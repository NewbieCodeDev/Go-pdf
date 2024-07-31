package main

import (
	"fmt"
	"log"
	"math"

	"strings"

	"strconv"

	"github.com/jung-kurt/gofpdf"
)

func generateTableContent(pdf *gofpdf.Fpdf, dataArray [][]string, width []float64) {
	margin := 20.0
	headerHeight := 50.0
	footerHeight := 40.0

	pdf.SetFont("THSarabunNew", "", 14)

	// Calculate available height for content after header and before footer
	_, pageHeight := pdf.GetPageSize()

	contentEndY := pageHeight - footerHeight

	for _, row := range dataArray {
		//log.Println("row ",row)
		contentHeight := 10.0
		var lines []float64
		maxLineIndex := 0

		for i, data := range row {

			line := estimateLines(pdf, data, width[i])
			lines = append(lines, line)

			if line > lines[maxLineIndex] {

				maxLineIndex = i

			}
		}

		// Check if content exceeds remaining height and adjust if necessary
		for i, data := range row {
			_, y := pdf.GetXY()

			remainingHeight := contentEndY - y
			//log.Println("remainingHeight ", remainingHeight)
			if contentHeight*lines[maxLineIndex] > remainingHeight {
				//log.Println("if")
				pdf.AddPage()
				pdf.SetY(margin + headerHeight)
				_, y = pdf.GetXY()
				remainingHeight = contentEndY - y
			}
			currentX, currentY := pdf.GetXY()

			if i == maxLineIndex {
				contentHeight = 10.0
				pdf.MultiCell(width[i], contentHeight, data, "1", "L", false)
			} else {

				pdf.MultiCell(width[i], 10.0*(lines[maxLineIndex]/lines[i]), data, "1", "L", false)
			}

			pdf.SetXY(currentX+width[i], currentY)

		}
		pdf.Ln(contentHeight * lines[maxLineIndex])

	}

}

func estimateLines(pdf *gofpdf.Fpdf, text string, maxWidth float64) float64 {

	lines := 0.0
	log.Println(pdf.GetFontSize())

	textList := strings.Split(text, "\n")
	for _, text := range textList {
		textWidth := pdf.GetStringWidth(text)
		lines += math.Ceil(textWidth / maxWidth)

	}

	log.Println(text)
	log.Println(lines)
	return float64(lines)
}

func generateTextContent(pdf *gofpdf.Fpdf, tabStartParagraph bool, content string) {
	margin := 20.0
	headerHeight := 50.0
	footerHeight := 40.0
	pageWidth, _, _ := pdf.PageSize(pdf.PageNo())
	currentY := pdf.GetY()
	//log.Println(currentY)

	pdf.SetFont("THSarabunNew", "", 14)
	indent := "           " // This is equivalent to a tab
	var indentedContent string

	// Calculate available height for content after header and before footer
	_, pageHeight := pdf.GetPageSize()

	contentEndY := pageHeight - footerHeight
	remainingHeight := contentEndY - currentY

	line1 := estimateLines(pdf, content, pageWidth-margin)
	contentHeight := line1 * pdf.PointConvert(14)

	if tabStartParagraph {
		indentedContent = indent + content
	} else {
		indentedContent = content
	}
	// Check if content exceeds remaining height and adjust if necessary
	if contentHeight > remainingHeight {
		// Content exceeds remaining height, handle overflow

		pdf.AddPage()
		pdf.SetY(margin + headerHeight)
		pdf.MultiCell(190, 10, indentedContent, "", "L", false)

	} else {

		// Content fits within remaining height
		pdf.SetY(currentY)
		pdf.MultiCell(190, 10, indentedContent, "", "L", false)

	}

}

func generateImageContent(pdf *gofpdf.Fpdf, imgList []string, w float64, h float64, margin float64, column bool) {

	pdf.Ln(5)
	// The width is 100 mm (as set), height is automatically calculated

	for i, path := range imgList {
		//log.Print("path", path)

		x := margin + float64(i)*(w+margin)

		y := pdf.GetY()
		pdf.Image(path, x, y, w, h, false, "", 0, "")

		if column {
			//log.Println("y", y)
			y := pdf.GetY()
			//log.Println("y", y)
			// Retrieve the image information
			info := pdf.GetImageInfo(path)

			// The width is 100 mm (as set), height is automatically calculated
			calculatedHeight := info.Height() * (100 / info.Width())
			pdf.SetY(y + calculatedHeight + margin + 10)
		}

	}
	//pdf.Ln(h + 20)
}

func main() {

	// Create a new PDF document
	pdf := gofpdf.New("P", "mm", "A4", "")
	fontSize := 20.0

	// Load a Thai font
	pdf.AddUTF8Font("THSarabunNew", "", "./THSarabunNew.ttf")
	pdf.AddUTF8Font("THSarabunNew", "B", "./THSarabunNew Bold.ttf")

	// Set font
	pdf.SetFont("THSarabunNew", "B", 28)

	generateHeader := func() {
		if pdf.PageNo() > 1 {
			pdf.SetFont("THSarabunNew", "B", 16)
			pdf.CellFormat(25, 15, "", "TL", 0, "L", false, 0, "")
			pdf.CellFormat(110, 15, "รายงานการปล่อยและดูดกลับก๊าซเรือนกระจก", "1", 0, "C", false, 0, "")
			pdf.SetFont("THSarabunNew", "B", 12)
			pdf.MultiCell(25, 5, "TCFO_R_02 Version 03.00 24/4/2019", "TRB", "L", false)

			pdf.SetFont("THSarabunNew", "", 12)
			pdf.CellFormat(25, 15, "", "L", 0, "L", false, 0, "")
			pdf.CellFormat(25, 15, "องค์กร", "1", 0, "L", false, 0, "")
			x, y := pdf.GetXY()
			pdf.MultiCell(85, 7.5, "บริษัท เบทาโกร จำกัด (มหาชน) โรงงานผลิตอาหารสัตว์ จ.ลพบุรี 1,2 และ 3", "1", "L", false)
			pdf.SetXY(x+85, y)
			pdf.CellFormat(25, 15, "หน้าที่ "+strconv.Itoa(pdf.PageNo()-1), "TR", 0, "L", false, 0, "")
			pdf.Ln(-1)

			pdf.SetFont("THSarabunNew", "", 12)
			pdf.CellFormat(25, 5, "", "LB", 0, "L", false, 0, "")
			pdf.CellFormat(25, 5, "หน่วยงานสอบทาน", "1", 0, "L", false, 0, "")
			x, y = pdf.GetXY()
			pdf.MultiCell(85, 5, "บริษัท อีซีอีอี จำกัด", "1", "L", false)
			pdf.SetXY(x+85, y)
			pdf.CellFormat(25, 5, "", "TRB", 0, "L", false, 0, "")

			pdf.Ln(20)
		}
	}

	generateFooter := func() {
		// Footer
		if pdf.PageNo() > 1 {
			pdf.SetY(-30)
			// pdf.SetFont("Arial", "", 8)
			pdf.SetFont("THSarabunNew", "B", 14)
			pdf.CellFormat(25, 7, "จัดทำโดย", "1", 0, "C", false, 0, "")
			pdf.CellFormat(65, 7, "สภาอุตสาหกรรมแห่งประเทศไทย", "1", 0, "L", false, 0, "")
			pdf.CellFormat(25, 7, "ผู้ทวนสอบ", "1", 0, "C", false, 0, "")
			pdf.CellFormat(50, 7, "บริษัท อีซีอีอี จำกัด", "1", 0, "C", false, 0, "")
		}
	}

	// Set header and footer functions
	pdf.SetHeaderFunc(generateHeader)
	pdf.SetFooterFunc(generateFooter)
	margin := 50.0
	pdf.SetMargins(margin, margin, margin)
	pdf.SetAutoPageBreak(true, margin)

	// Add a page
	pdf.AddPage()

	// Write Thai text
	// Add title

	pdf.CellFormat(0, 10, "รายงานการปล่อยและดูดกลับก๊าซเรือนกระจก", "", 2, "C", false, 0, "")
	pdf.Ln(10)

	// Add images
	imagePaths := []string{"rabbit.jpg", "rabbit.jpg", "rabbit.jpg", "rabbit.jpg"}
	generateImageContent(pdf, imagePaths, 45.0, 45.0, 5.0, false)
	pdf.Ln(60)

	pdf.SetFont("THSarabunNew", "B", fontSize)
	pdf.CellFormat(25, 10, "ชื่อองค์กร : ", "", 0, "L", false, 0, "")
	pdf.SetFont("THSarabunNew", "", fontSize)
	pdf.MultiCell(0, 10, "บริษัท เบทาโกร จำกัด (มหาชน) โรงงานผลิตอาหารสัตว์ จ.ลพบุรี 1,2 และ 3", "", "L", false)

	pdf.SetFont("THSarabunNew", "B", fontSize)
	pdf.CellFormat(50, 10, "ที่อยู่/สถานที่ตั้งองค์กร : ", "", 0, "L", false, 0, "")
	pdf.SetFont("THSarabunNew", "", fontSize)
	pdf.MultiCell(0, 10, "เลขที่ 3 หมู่ 13 ถ.สระบุรี-หล่มสัก ต.ช่องสาริกา อ.พัฒนานิคม จ.ลพบุรี", "", "L", false)

	pdf.SetFont("THSarabunNew", "B", fontSize)
	pdf.CellFormat(35, 10, "วันที่รายงานผล : ", "", 0, "L", false, 0, "")
	pdf.SetFont("THSarabunNew", "", fontSize)
	pdf.CellFormat(0, 10, "28 มิ.ย. 2566", "", 1, "L", false, 0, "")

	pdf.SetFont("THSarabunNew", "B", fontSize)
	pdf.CellFormat(55, 10, "ระยะเวลาในการติดตามผล : ", "", 0, "L", false, 0, "")
	pdf.SetFont("THSarabunNew", "", fontSize)
	pdf.CellFormat(0, 10, "มกราคม ถึง ธันวาคม 2565", "", 1, "L", false, 0, "")
	pdf.Ln(20)

	pdf.SetFont("THSarabunNew", "B", 18)

	// Add footer
	pdf.CellFormat(0, 0, "เพื่อการทวนสอบและรับรองผลคาร์บอนฟุตพริ้นท์ขององค์กร", "", 2, "C", false, 0, "")
	pdf.Ln(10)
	pdf.CellFormat(0, 0, "โดย องค์การบริหารจัดการก๊าซเรือนกระจก (องค์การมหาชน)", "", 2, "C", false, 0, "")
	pdf.Ln(50)

	// Page 1
	margin = 20.0
	pdf.SetMargins(margin, margin, margin)
	pdf.SetAutoPageBreak(true, margin)

	pdf.AddPage()

	// Header
	// Table header

	// Title
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "1. บทนำ", "", 1, "L", false, 0, "")

	// Add some space before the paragraph
	pdf.CellFormat(0, 8, "", "0", 1, "C", false, 0, "")

	// Content
	// paragraph 1
	content := `จากผลกระทบของภาวะโลกร้อน ทำให้ประเทศต่างๆ ทั่วโลกตื่นตัวในการดำเนินงานเพื่อลดการปล่อยก๊าซเรือนกระจก แนวคิดการจัดทำคาร์บอนฟุตพริ้นท์ขององค์กร (Carbon Footprint for Organization: CFO) เป็นวิธีการประเมินปริมาณก๊าซเรือนกระจกที่ปล่อยจากกิจกรรมทั้งหมดขององค์กรและคำนวณออกมาในรูปคาร์บอนไดออกไซด์เทียบเท่า อันจะนำไปสู่การกำหนดแนวทางการบริหารจัดการ เพื่อลดการปล่อยก๊าซเรือนกระจกได้อย่างมีประสิทธิภาพทั้งในระดับหน่วยงาน บริษัท หรือโรงงาน ระดับอุตสาหกรรม และระดับประเทศ `
	generateTextContent(pdf, true, content)

	// paragraph 2
	content = `จากผลกระทบของภาวะโลกร้อน ทำให้ประเทศต่างๆ ทั่วโลกตื่นตัวในการดำเนินงานเพื่อลดการปล่อยก๊าซเรือนกระจก แนวคิดการจัดทำคาร์บอนฟุตพริ้นท์ขององค์กร (Carbon Footprint for Organization: CFO) เป็นวิธีการประเมินปริมาณก๊าซเรือนกระจกที่ปล่อยจากกิจกรรมทั้งหมดขององค์กรและคำนวณออกมาในรูปคาร์บอนไดออกไซด์เทียบเท่า อันจะนำไปสู่การกำหนดแนวทางการบริหารจัดการ เพื่อลดการปล่อยก๊าซเรือนกระจกได้อย่างมีประสิทธิภาพทั้งในระดับหน่วยงาน บริษัท หรือโรงงาน ระดับอุตสาหกรรม และระดับประเทศ `
	generateTextContent(pdf, true, content)

	// paragraph 3
	content = `จากผลกระทบของภาวะโลกร้อน ทำให้ประเทศต่างๆ ทั่วโลกตื่นตัวในการดำเนินงานเพื่อลดการปล่อยก๊าซเรือนกระจก แนวคิดการจัดทำคาร์บอนฟุตพริ้นท์ขององค์กร (Carbon Footprint for Organization: CFO) เป็นวิธีการประเมินปริมาณก๊าซเรือนกระจกที่ปล่อยจากกิจกรรมทั้งหมดขององค์กรและคำนวณออกมาในรูปคาร์บอนไดออกไซด์เทียบเท่า อันจะนำไปสู่การกำหนดแนวทางการบริหารจัดการ เพื่อลดการปล่อยก๊าซเรือนกระจกได้อย่างมีประสิทธิภาพทั้งในระดับหน่วยงาน บริษัท หรือโรงงาน ระดับอุตสาหกรรม และระดับประเทศ `
	generateTextContent(pdf, true, content)

	//page 2

	content = `กรกฎาคม 2565) ขององค์การบริหารจัดการก๊าซเรือนกระจก (องค์การมหาชน) และขอรับการทวนสอบข้อมูลเป็นระดับการทวนสอบแบบจำกัด (Limited level of assurance) และมีความมีสาระสำคัญ(Materiality) 5% `
	generateTextContent(pdf, false, content)

	pdf.SetFont("THSarabunNew", "B", 16)
	// Add some space before the paragraph
	pdf.CellFormat(0, 8, "", "0", 1, "C", false, 0, "")
	pdf.AddPage()
	pdf.CellFormat(0, 10, "2. ข้อมูลทั่วไป ", "", 1, "L", false, 0, "")
	// Add some space before the paragraph
	pdf.CellFormat(0, 8, "", "0", 1, "C", false, 0, "")

	//content
	pdf.SetFont("THSarabunNew", "B", 14)
	data := [][]string{
		{"2.1 ชื่อองค์กร ", "บริษัท เบทาโกร จำกัด (มหาชน) โรงงานผลิตอาหารสัตว์ จ.ลพบุรี 1,2 และ 3 "},
		{"2.2 ที่อยู่/สถานที่ตั้งองค์กร ", "เลขที่ 3 หมู่ 13 ถ.สระบุรี-หล่มสัก ต.ช่องสาริกา อ.พัฒนานิคม จ.ลพบุรี"},
		{"2.3 ประเภทของอุตสาหกรรม ", "ผู้ผลิตอาหารสัตว์ "},
		{"2.4 ชื่อ-สกุลของผู้ประสานงาน ", "1.คุณวนิตา ทัลวัลลิ์ \n2.คุณสุวรรณา แก้วกล่ำ "},
		{"2.5 ชื่อ-สกุลของผู้รับผิดชอบข้อมูล ", "1.คุณเบญจมา กลีบทอง \n2.คุณจีรประภา วงษ์พาศกลาง \n3.คุณปพิภากาญจณ์ สุวรรณวงษ์"},
		{"2.6 ระยะเวลาติดตามผล ", "มกราคม ถึง ธันวาคม 2565 "},
		{"2.7 แนวทางที่ใช้ในการติดตามผล ", "ข้อกำหนดในการคำนวณและรายงานคาร์บอนฟุตพริ้นท์ขององค์กร พิมพ์ครั้งที่ 8 (ฉบับปรับปรุงครั้งที่ 6 กรกฎาคม 2565)  "},
		{"2.8 ระดับของการรับรอง (Level of Assurance)", "แบบจำกัด (Limited Assurance)"},
		{"2.9 ระดับความมีสาระสำคัญ (Materiality Threshold)  ", "5% Materiality"},
	}

	//generateTableContent(pdf, data, []float64{60.0, 120.0})

	//end page 2

	//page 3
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "3. ขอบเขต  ", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, " 3.1 ขอบเขตขององค์กร ", "", 1, "L", false, 0, "")

	data = [][]string{
		{"1) แนวทางที่ใช้กำหนดขอบเขตองค์กร ", "ควบคุมดำเนินงาน (OPERATIONAL CONTROL) "},
		{"2) หน่วยสาธารณูปโภค (Facility)/พื้นที่ที่ครอบคลุมในรายงาน ", "1. บริษัท เบทาโกร จำกัด (มหาชน)  โรงงานลพบุรี 1 \n2. บริษัท เบทาโกร จำกัด (มหาชน)  โรงงานลพบุรี 2 \n3. บริษัท เบทาโกร จำกัด (มหาชน)  โรงงานลพบุรี 3"},
		{"3) เอกสารยืนยันขอบเขต ", "โรงงานลพบุรี 1,2,3 : ใบอนุญาตประกอบกิจการโรงงานเลขที ่ :  ส3-15(1)-1/34ลบ "},
	}

	//generateTableContent(pdf, data, []float64{60.0, 120.0})
	// end page 3

	//page 4
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "3.1.1 โครงสร้างขององค์กร", "", 1, "L", false, 0, "")

	// Add images
	imagePaths = []string{"companyStructure.png"}
	generateImageContent(pdf, imagePaths, 150.0, 0.0, 15.0, true)

	// end page 4

	//page 5
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "3.1.2 แผนผังของโรงงาน ", "", 1, "L", false, 0, "")

	// Add images
	imagePaths = []string{"companyMap.png"}
	generateImageContent(pdf, imagePaths, 0.0, 180.0, 15.0, true)

	// end page 5

	//page 6
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "3.1.3 แผนผังกระบวนการผลิต ", "", 1, "L", false, 0, "")

	// Add images
	imagePaths = []string{"productionMap.png"}
	generateImageContent(pdf, imagePaths, 150.0, 0.0, 15.0, true)
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.CellFormat(0, 10, "รูปแสดง : แผนผังการผลิต บริษัท เบทาโกร จำกัด (มหาชน)  โรงงานลพบุรี 1", "", 1, "C", false, 0, "")

	// Add images
	imagePaths = []string{"productionMap1.png"}
	generateImageContent(pdf, imagePaths, 150.0, 0.0, 15.0, true)
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.CellFormat(0, 10, "รูปแสดง : แผนผังการผลิต บริษัท เบทาโกร จำกัด (มหาชน)  โรงงานลพบุรี 2", "", 1, "C", false, 0, "")

	// end page 6

	//page 7
	pdf.AddPage()
	// Add images
	imagePaths = []string{"productionMap2.png"}
	generateImageContent(pdf, imagePaths, 150.0, 0.0, 15.0, true)
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.CellFormat(0, 10, "รูปแสดง : แผนผังการผลิต บริษัท เบทาโกร จำกัด (มหาชน)  โรงงานลพบุรี 3", "", 1, "C", false, 0, "")

	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "3.1.4 ระบุกิจกรรมทั้งหมดขององค์กร  ", "", 1, "L", false, 0, "")

	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.CellFormat(0, 10, "จำแนกกิจกรรมขององค์กรในแต่ละ Facility  (ใส่หมายเลขและชื่อ Facility ในข้อ 3.1.2) ตามแผนผังให้ครอบคลุมทุก Scopes ", "", 1, "L", false, 0, "")

	// Set text color to black
	pdf.SetTextColor(0, 0, 0)
	data = [][]string{
		{"Facility ", "กิจกรรมขององค์กรในแต่ละ Facility ", "", ""},
		{"", "Scope 1 ", "Scope 2 ", "Scope 3"},
		{"1. บริษัท เบทาโกร จำกัด (มหาชน)  โรงงานลพบุรี 1 (LR1)", "1. การเผาไหม้น้ำมันดีเซลรถยนต์ ", "1. การใช้ไฟฟ้า ", "1. Purchased goods and services "},
		{"", "2. การเผาไหม้น้ำมันเบนซีนรถยนต์ ", "", "2. Fuel- and energy related activities not included scope 1 & 2"},
		{"", "3. การเผาไหม้ก๊าซ LPG สำหรับรถยนต์", "", "3. Upstream transport "},
		{"", "4. การเผาไหม้ก๊าซ NGV สำหรับรถยนต", "", "4. Waste generate"},
		{"", "5. การเผาไหม้น้ำมันดีเซล generator + fire pump ", "", "5. Downstream transport"},
		{"", "6. การเผาไหม้น้ำมันดีเซลเครื่องตัดหญ้า ", "", ""},
		{"", "7. การเผาไหม้น้ำมันเตา C Boiler และการอบข้าวโพด", "", "4. Waste generate"},
	}
	pdf.SetFont("THSarabunNew", "", 12)
	// Set fill color (RGB)
	pdf.SetFillColor(190, 190, 190)
	pdf.CellFormat(50, 15, "Facility ", "TL", 0, "C", true, 0, "")
	pdf.CellFormat(120, 15, "กิจกรรมขององค์กรในแต่ละ Facility ", "1", 0, "C", true, 0, "")
	pdf.Ln(-1)

	pdf.CellFormat(50, 7.5, "", "LB", 0, "L", true, 0, "")
	pdf.CellFormat(50, 7.5, "Scope 1 ", "1", 0, "C", true, 0, "")
	x, y := pdf.GetXY()
	pdf.MultiCell(25, 7.5, "Scope 2 ", "1", "C", true)
	pdf.SetXY(x+25, y)
	pdf.CellFormat(45, 7.5, "Scope 3", "1", 0, "C", true, 0, "")
	pdf.Ln(-1)

	// Set fill color (RGB)
	pdf.SetFillColor(255, 255, 255)

	data = [][]string{
		{"1. บริษัท เบทาโกร จำกัด (มหาชน) \nโรงงานลพบุรี 1 (LR1)", "1.การเผาไหม้น้ำมันดีเซลรถยนต", "1.การใช้ไฟฟ้า", " 1.Purchased goods and services"},
		{" ", "2.การเผาไหม้น้ำมันดีเซลรถยนต์", " ", " 2. Purchased goods and services"},
	}

	//generateTableContent(pdf, data, []float64{50.0, 50.0, 25.0, 45.0})

	//pdf.Ln(10)

	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ *กิจกรรมขององค์กรใน Scope 3 ที่ไม่รวมไว้ในการติดตามผล ", "", 0, "L", true, 0, "")

	//end page 7

	// end page 7

	//page 8
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.MultiCell(0, 10, "3.1.5 ระบุขอบเขตขององค์กรที่เพิ่มเข้ามาหรือขอบเขตที่ไม่รวม (ระบุ Facility) ที่เพิ่มเข้ามาหรือไม่นับรวม) พร้อมเหตุผล", "", "L", false)
	generateTextContent(pdf, true, "1. ไม่นับรวมการปล่อยก๊าซเรือนกระจกการใช้ก๊าซ LPG กิจกรรมซ่อมบำรุง โรงงานลพบุรี 1 ,2 เนื่องจากมีการใช้งานน้อยมาก มีอายุการใช้งานมากกว่า 1 ปี ")

	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.MultiCell(0, 10, "3.2 ขอบเขตการดำเนินงาน", "", "L", false)

	data = [][]string{
		{"1) ก๊าซเรือนกระจกที่พิจารณา ", "- คำร์บอนไดออกไซด์ (CO2)\n- มีเทน (CH4) \n- ไนตรัสออกไซด์ (N2O) \n- ไฮโดรฟลูออโรคำร์บอน (HFCs) \n- เพอร์ฟลูออโรคำร์บอน (PFCs) \n- ซัลเฟอร์เฮกซะฟลูออไรด์ (SF6) \n- ไนโตรเจนไตรฟลูออไรด์ (NF3)"},
		{"2) ก๊าซเรือนกระจกที่พิจารณาอื่น ๆเพิ่มเติม ", "-"},
		{"3) GWP ", "- IPCC Fifth Assessment Report (AR5)  "},
	}
	//generateTableContent(pdf, data, []float64{60.0, 120.0})
	//end page 8

	//page 9
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.MultiCell(0, 10, "3.2.1   ระบุกิจกรรมที่เป็นแหล่งปล่อยก๊าซเรือนกระจกประเภทที่1 ขององค์กร ", "", "L", false)

	pdf.SetFillColor(190, 190, 190)
	pdf.CellFormat(20, 15*4, "Facility ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "แหล่งปล่อยก๊าซเรือนกระจก (Emission Source) เช่น ระบุ อุปกรณ์หลัก/ เครื่องจักร / กระบวนการ/กิจกรรม ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 15*4, "ที่ตั้ง/ตำแหน่ง", "1", "C", true)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*4, "ใช้ภายใน ", "1", "C", true)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*2, "จำหน่ายภายนอก", "1", "C", true)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 20, "ความสำคัญ (มีนัยสำคัญมาก หรือ น้อย) ", "1", "C", true)

	pdf.SetFillColor(255, 255, 255)
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.CellFormat(170, 10, "Mobile Combustion ", "1", 0, "L", true, 0, "")
	pdf.Ln(-1)

	pdf.CellFormat(20, 15, "LR1,2,3  ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "1. น้ำมันดีเซลรถยนต์  ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 15, "-", "1", "C", true)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15, " ", "1", "C", true)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15, "", "1", "C", true)
	pdf.SetXY(x+20, y)
	pdf.MultiCell(30, 15, "น้อย", "1", "C", true)

	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.CellFormat(170, 10, "Stationary Combustion ", "1", 0, "L", true, 0, "")
	pdf.Ln(-1)

	pdf.CellFormat(20, 15, "LR1,2,3  ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "1. น้ำมันดีเซลรถยนต์  ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 15, "-", "1", "C", true)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15, " ", "1", "C", true)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15, "", "1", "C", true)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 15, "น้อย", "1", "C", true)
	pdf.Ln(-1)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ :  1. มีนัยสำคัญ “มาก” หมายถึง มีปริมาณการปล่อยก๊าซเรือนกระจกตั้งแต่ร้อยละ 5 ของปริมาณการปล่อยก๊าซเรือนกระจกรวมประเภทที่ 1+2 ขององค์กร", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "2. มีนัยสำคัญ “น้อย” หมายถึง มีปริมาณการปล่อยก๊าซเรือนกระจกน้อยกว่าร้อยละ 5 ของปริมาณการปล่อยก๊าซเรือนกระจกรวมประเภทที่ 1+2 ขององค์กร", "", 2, "L", false, 0, "")

	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.MultiCell(0, 10, "3.2.2 แหล่งปล่อยก๊าซเรือนกระจกทางตรงที่เกี่ยวข้องกับการใช้ชีวมวลและก๊าซชีวภาพ เพื่อทดแทนการใช้พลังงานและความร้อน  ", "", "L", false)

	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.MultiCell(0, 10, "พิจารณาเฉพาะที่มาจากพืช ของเสียอุตสาหกรรม และของเสียทั่วไป อ้างอิงตาม EB 23 Report Annex 18, DEFINITION OF RENEWABLE BIOMASS", "", "L", false)

	// Set text color to black
	pdf.SetTextColor(0, 0, 0)
	pdf.SetFillColor(190, 190, 190)
	pdf.CellFormat(20, 15*4, "Facility ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "แหล่งปล่อยก๊าซเรือนกระจก (Emission Source) เช่น ระบุ อุปกรณ์หลัก/ เครื่องจักร / กระบวนการ/กิจกรรม ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 15*4, "ที่ตั้ง/ตำแหน่ง", "1", "C", true)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*4, "ใช้ภายใน ", "1", "C", true)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*2, "จำหน่ายภายนอก", "1", "C", true)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 20, "ความสำคัญ (มีนัยสำคัญมาก หรือ น้อย) ", "1", "C", true)

	pdf.CellFormat(20, 10, " ", "1", 0, "C", false, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 10, "-ไม่มี-", "1", "C", false)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 10, "", "1", "C", false)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, " ", "1", "C", false)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, "", "1", "C", false)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 10, "", "1", "C", false)

	//end page 9

	//page10
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.MultiCell(0, 10, "3.2.3 ระบุกิจกรรมที่เป็นแหล่งปล่อยก๊าซเรือนกระจกทางตรงอื่น ๆ ที่ทำการรายงานแยก ", "", "L", false)

	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.MultiCell(0, 10, "ในกรณีที่มีการรายงานการปล่อยก๊ำซเรือนกระจกชนิดอื่น ๆ ที่ไม่อยู่ในข้อกำหนด เช่น R22 ให้ทำการรายงานแยก", "", "L", false)

	// Set text color to black
	pdf.SetTextColor(0, 0, 0)
	pdf.SetFillColor(190, 190, 190)
	pdf.CellFormat(20, 15*4, "Facility ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "แหล่งปล่อยก๊าซเรือนกระจก (Emission Source) เช่น ระบุ อุปกรณ์หลัก/ เครื่องจักร / กระบวนการ/กิจกรรม ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 15*4, "ที่ตั้ง/ตำแหน่ง", "1", "C", true)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*4, "ใช้ภายใน ", "1", "C", true)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*2, "จำหน่ายภายนอก", "1", "C", true)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 20, "ความสำคัญ (มีนัยสำคัญมาก หรือ น้อย) ", "1", "C", true)

	pdf.CellFormat(20, 10, " ", "1", 0, "C", false, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 10, "-ไม่มี-", "1", "C", false)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 10, "", "1", "C", false)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, " ", "1", "C", false)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, "", "1", "C", false)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 10, "", "1", "C", false)
	pdf.Ln(-1)

	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.MultiCell(0, 10, "3.2.4 ระบุกิจกรรมที่เป็นแหล่งปล่อยก๊าซเรือนกระจกประเภทที่ 2 ขององค์กร", "", "L", false)

	pdf.SetFillColor(190, 190, 190)
	pdf.CellFormat(20, 15*4, "Facility ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "แหล่งปล่อยก๊าซเรือนกระจก (Emission Source) เช่น ระบุ อุปกรณ์หลัก/ เครื่องจักร / กระบวนการ/กิจกรรม ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 15*4, "ที่ตั้ง/ตำแหน่ง", "1", "C", true)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*4, "ใช้ภายใน ", "1", "C", true)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*2, "จำหน่ายภายนอก", "1", "C", true)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 20, "ความสำคัญ (มีนัยสำคัญมาก หรือ น้อย) ", "1", "C", true)

	pdf.CellFormat(20, 10, " ", "1", 0, "C", false, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 10, "-ไม่มี-", "1", "C", false)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 10, "", "1", "C", false)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, " ", "1", "C", false)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, "", "1", "C", false)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 10, "", "1", "C", false)

	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ :  1. มีนัยสำคัญ “มาก” หมายถึง มีปริมาณการปล่อยก๊าซเรือนกระจกตั้งแต่ร้อยละ 5 ของปริมาณการปล่อยก๊าซเรือนกระจกรวมประเภทที่ 1+2 ขององค์กร", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "2. มีนัยสำคัญ “น้อย” หมายถึง มีปริมาณการปล่อยก๊าซเรือนกระจกน้อยกว่าร้อยละ 5 ของปริมาณการปล่อยก๊าซเรือนกระจกรวมประเภทที่ 1+2 ขององค์กร", "", 2, "L", false, 0, "")

	// 3.2.5 table
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 14)
	pdf.MultiCell(0, 10, "3.2.5 พลังงาน/ความร้อน/ไอน้ำที่จำหน่ายให้หน่วยงานภายนอก (Supply to External) (นอกขอบเขตการดำเนินงาน) (out of boundary) ", "", "L", false)
	data = [][]string{
		{"อุปกรณ์ / เครื่องจักรที่ผลิตพลังงาน / ความร้อน / ไอน้ำ / กระบวนการ (Source)  ", "จำหน่ายให้กับ (Supply to) "},
		{"-ไม่มี-", " "},
	}
	//generateTableContent(pdf, data, []float64{60.0, 120.0})

	//3.2.6 table
	pdf.MultiCell(0, 10, "3.2.6   ระบุกิจกรรมที่เป็นแหล่งปล่อยก๊าซเรือนกระจกประเภทที่ 3 ขององค์กร", "", "L", false)
	pdf.SetFillColor(190, 190, 190)
	pdf.CellFormat(20, 15*4, "Facility ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "แหล่งปล่อยก๊าซเรือนกระจก (Emission Source) เช่น ระบุ อุปกรณ์หลัก/ เครื่องจักร / กระบวนการ/กิจกรรม ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 15*4, "ที่ตั้ง/ตำแหน่ง", "1", "C", true)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*4, "ใช้ภายใน ", "1", "C", true)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 15*2, "จำหน่ายภายนอก", "1", "C", true)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 20, "ความสำคัญ (มีนัยสำคัญมาก หรือ น้อย) ", "1", "C", true)

	pdf.CellFormat(20, 10, " ", "1", 0, "C", false, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 10, "-ไม่มี-", "1", "C", false)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 10, "", "1", "C", false)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, " ", "1", "C", false)
	pdf.SetXY(x+20, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(20, 10, "", "1", "C", false)
	pdf.SetXY(x+20, y)

	pdf.MultiCell(30, 10, "", "1", "C", false)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ :  1. มีนัยสำคัญ “มาก” หมายถึง มีปริมาณการปล่อยก๊าซเรือนกระจกตั้งแต่ร้อยละ 5 ของปริมาณการปล่อยก๊าซเรือนกระจกรวมประเภทที่ 1+2 ขององค์กร", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "2. มีนัยสำคัญ “น้อย” หมายถึง มีปริมาณการปล่อยก๊าซเรือนกระจกน้อยกว่าร้อยละ 5 ของปริมาณการปล่อยก๊าซเรือนกระจกรวมประเภทที่ 1+2 ขององค์กร", "", 2, "L", false, 0, "")

	//3.2.7 table
	pdf.AddPage()
	pdf.MultiCell(0, 10, "3.2.7 การกักเก็บคาร์บอน", "", "L", false)
	pdf.SetFillColor(190, 190, 190)

	x, y = pdf.GetXY()
	pdf.MultiCell(40, 15, "รายชื่อกระบวนการ (Sink / Reservoir) ", "1", "C", true)
	pdf.SetXY(x+40, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(40, 15, "กำลังการผลิต (Capacity)", "1", "C", true)
	pdf.SetXY(x+40, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(40, 15, "ที่ตั้ง/ตำแหน่ง", "1", "C", true)
	pdf.SetXY(x+40, y)

	pdf.MultiCell(40, 15, "ความสำคัญ (มีนัยสำคัญมาก หรือ น้อย) ", "1", "C", true)

	pdf.CellFormat(40, 10, " -ไม่มี-", "1", 0, "C", false, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(40, 10, "", "1", "C", false)
	pdf.SetXY(x+40, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(40, 10, "", "1", "C", false)
	pdf.SetXY(x+40, y)

	pdf.MultiCell(40, 10, " ", "1", "C", false)
	pdf.Ln(-1)

	//3.2.8 table
	//pdf.AddPage()
	pdf.MultiCell(0, 10, "3.2.8 โครงการลดก๊าซเรือนกระจก/การรับรองสิทธิพลังงานหมุนเวียน ", "", "L", false)
	pdf.SetFillColor(190, 190, 190)

	x, y = pdf.GetXY()
	pdf.MultiCell(40, 15, "ชื่อโครงการ", "1", "C", true)
	pdf.SetXY(x+40, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(40, 15, "มาตรฐานที่ของรับรอง", "1", "C", true)
	pdf.SetXY(x+40, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(40, 15, "ระยะเวลาคิดคาร์บอนเครดิตของโครงการ", "1", "C", true)
	pdf.SetXY(x+40, y)

	pdf.MultiCell(40, 5, "จำนวนคาร์บอนเครดิต/สิทธิพลังงานหมุนเวียน ที ่ได้รับการรับรองที ่ขายไป (TonCO2e/kWh) ", "1", "C", true)

	pdf.CellFormat(40, 10, " -ไม่มี-", "1", 0, "C", false, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(40, 10, "", "1", "C", false)
	pdf.SetXY(x+40, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(40, 10, "", "1", "C", false)
	pdf.SetXY(x+40, y)

	pdf.MultiCell(40, 10, " ", "1", "C", false)

	// 4.
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "4. การติดตามผล  ", "", 1, "L", false, 0, "")
	// Set text color to black
	pdf.SetTextColor(255, 0, 0)
	pdf.MultiCell(0, 10, "จุดที่ตรวจวัด หมายถึง ตำแหน่งมิเตอร์ (อ้างอิงแผนผังมิเตอร์หรืออุปกรณ์ตรวจวัด ในภาคผนวก 1) หรือ จุดที่มีการบันทึกข้อมูล (อ้างอิงตามโครงสร้างระบบการจัดการคุณภาพของข้อมูลในข้อ 7.1) ", "", "L", false)
	// Set text color to black
	pdf.SetTextColor(0, 0, 0)
	pdf.CellFormat(0, 10, "4.1 แหล่งปล่อยก๊าซเรือนกระจก จากขอบเขตการดำเนินงานประเภทที่ 1", "", 1, "L", false, 0, "")

	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(35, 7.5, "", "TL", 0, "L", true, 0, "")
	pdf.CellFormat(110, 7.5, "ข้อมูลกิจกรรม", "1", 0, "C", true, 0, "")
	pdf.SetFont("THSarabunNew", "B", 12)
	pdf.MultiCell(10, 7.5, "ค่า EF", "1", "L", true)

	pdf.SetFont("THSarabunNew", "", 12)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 22.5, "แหล่งปล่อยก๊าซเรือนกระจก ", "BL", "C", true)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 7.5, "ลักษณะข้อมูลกิจกรรมที่ตรวจวัด ", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 11.25, "จุดที่\nตรวจวัด", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(54, 7.5, "ที่มาของข้อมูลกิจกรรม ", "1", "C", true)
	pdf.SetXY(x, y+7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 7.5, "เป็นค่าที่ได้จากการตรวจวัด", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากหลักฐานการชำระเงิน", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากการประเมินค่า", "1", "L", true)
	pdf.SetXY(x+18, y-7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 11.25, "หลักฐาน/\nเอกสารอ้างอิง", "1", "L", true)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 7.5, "ที่มา\nของค่า EF", "1", "L", true)

	pdf.CellFormat(155, 7.5, "Mobile Combustion ", "1", 0, "L", false, 0, "")
	pdf.Ln(-1)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 28.35, "น้ำมันดีเซลรถยนต์ ", "BL", "C", false)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 11.4, "ลิตร", "1", "L", false)
	pdf.SetXY(x, y+11.4)
	pdf.MultiCell(15, 17.1, "บาท", "1", "L", false)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 5.7, "แผนก\nยานยนต์ ", "1", "L", false)
	pdf.SetXY(x, y+11.4)
	pdf.MultiCell(15, 5.7, "แผนก\nทรัพยากร\nมนุษย์ ", "1", "L", false)
	pdf.SetXY(x+15, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(18, 11.4, "", "1", "L", false)
	pdf.SetXY(x, y+11.4)
	pdf.MultiCell(18, 17.1, "", "1", "L", false)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 11.4, "", "1", "L", false)
	pdf.SetXY(x, y+11.4)
	pdf.MultiCell(18, 17.1, "", "1", "L", false)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 11.4, "", "1", "L", false)
	pdf.SetXY(x, y+11.4)
	pdf.MultiCell(18, 17.1, "", "1", "L", false)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 5.7, "รายงานสรุป Fleet card ", "1", "L", false)
	pdf.SetXY(x, y+11.4)
	pdf.MultiCell(26, 4.275, "1. ยอดเบิกเงินจาก SAP \n2. ราคาน้ำมัน\nเฉลี่ยรายเดือน", "1", "L", false)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 9.45, "CFO\nTGO\nEF ", "1", "L", false)

	pdf.Ln(-1)
	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(1)  ข้อมูลกิจกรรมที ่ได้จากการตรวจวัด ให้ระบุรายละเอียดการสอบเทียบของอุปกรณ์ตรวจวัดไว้ในตารางที ่ 7.3  ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(2) ข้อมูลกิจกรรมที่ได้จากการประมาณค่า ให้อธิบายแนวทางในการประมาณในตารางหรืออธิบายเพิ่มเติมในภาคผนวก  ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(3) ในกรณีที่ข้อมูลกิจกรรมเป็นข้อมูลปริมาณการปล่อยก๊าซเรือนกระจกอยู่แล้ว เช่น ปริมาณการรั ่วซึมของสารทำความเย็น ให้กรอกคำว่า “ไม่ต้องใช้ค่า EF” ลงในคอลัมน์ “ที่มาของค่า EF” ", "", 2, "L", false, 0, "")

	// 4.2
	// Set text color to black

	pdf.SetTextColor(0, 0, 0)
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, "4.2 แหล่งปล่อยก๊าซเรือนกระจก จากขอบเขตการดำเนินงานประเภทที่ 2 ", "", 1, "L", false, 0, "")

	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(35, 7.5, "", "TL", 0, "L", true, 0, "")
	pdf.CellFormat(110, 7.5, "ข้อมูลกิจกรรม", "1", 0, "C", true, 0, "")
	pdf.SetFont("THSarabunNew", "B", 12)
	pdf.MultiCell(10, 7.5, "ค่า EF", "1", "L", true)

	pdf.SetFont("THSarabunNew", "", 12)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 22.5, "แหล่งปล่อยก๊าซเรือนกระจก ", "BL", "C", true)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 7.5, "ลักษณะข้อมูลกิจกรรมที่ตรวจวัด ", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 11.25, "จุดที่\nตรวจวัด", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(54, 7.5, "ที่มาของข้อมูลกิจกรรม ", "1", "C", true)
	pdf.SetXY(x, y+7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 7.5, "เป็นค่าที่ได้จากการตรวจวัด", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากหลักฐานการชำระเงิน", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากการประเมินค่า", "1", "L", true)
	pdf.SetXY(x+18, y-7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 11.25, "หลักฐาน/\nเอกสารอ้างอิง", "1", "L", true)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 7.5, "ที่มา\nของค่า EF", "1", "L", true)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 28.35, "การใช้ไฟฟ้า", "BL", "C", false)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 28.35, "kWh", "1", "C", false)

	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 14.175, "แผนก\nพลังงาน ", "1", "L", false)
	pdf.SetXY(x+15, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(18, 28.35, "", "1", "L", false)

	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 28.35, "", "1", "L", false)

	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 28.35, "", "1", "L", false)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 4.05, "1. รายงานการ\nใช้ไฟฟ้า (จริง) rate 115 kv ประจำเดือน\nของโรงงาน  \n2. หนังสือแจ้ง\nค่าไฟฟ้าจากการ\nไฟฟ้าส่วนภูมิภาค ", "1", "L", false)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 9.45, "CFO\nTGO\nEF ", "1", "L", false)

	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(1)  ข้อมูลกิจกรรมที ่ได้จากการตรวจวัด ให้ระบุรายละเอียดการสอบเทียบของอุปกรณ์ตรวจวัดไว้ในตารางที ่ 7.3   ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(2) ข้อมูลกิจกรรมที่ได้จากการประมาณค่า ให้อธิบายแนวทางในการประมาณในตารางหรืออธิบายเพิ่มเติมในภาคผนวก  ", "", 2, "L", false, 0, "")

	// 4.3
	// Set text color to black

	pdf.SetTextColor(0, 0, 0)
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, "4.3 แหล่งปล่อยก๊าซเรือนกระจก จากขอบเขตการด าเนินงานประเภทที ่ 3", "", 1, "L", false, 0, "")

	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(35, 7.5, "", "TL", 0, "L", true, 0, "")
	pdf.CellFormat(110, 7.5, "ข้อมูลกิจกรรม", "1", 0, "C", true, 0, "")
	pdf.SetFont("THSarabunNew", "B", 12)
	pdf.MultiCell(10, 7.5, "ค่า EF", "1", "L", true)

	pdf.SetFont("THSarabunNew", "", 12)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 22.5, "แหล่งปล่อยก๊าซเรือนกระจก ", "BL", "C", true)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 7.5, "ลักษณะข้อมูลกิจกรรมที่ตรวจวัด ", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 11.25, "จุดที่\nตรวจวัด", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(54, 7.5, "ที่มาของข้อมูลกิจกรรม ", "1", "C", true)
	pdf.SetXY(x, y+7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 7.5, "เป็นค่าที่ได้จากการตรวจวัด", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากหลักฐานการชำระเงิน", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากการประเมินค่า", "1", "L", true)
	pdf.SetXY(x+18, y-7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 11.25, "หลักฐาน/\nเอกสารอ้างอิง", "1", "L", true)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 7.5, "ที่มา\nของค่า EF", "1", "L", true)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 14.175, "1. Purchased goods \nand services ", "BL", "C", false)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 28.35, "กก.", "1", "C", false)

	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 14.175, "แผนกคลัง\nวัตถุดิบ", "1", "L", false)
	pdf.SetXY(x+15, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(18, 28.35, "", "1", "L", false)

	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 28.35, "", "1", "L", false)

	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 28.35, "", "1", "L", false)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 7.05, "1. ข้อมูลการ\nรับเข้าจาก\nระบบ \nSAP  ", "1", "L", false)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 9.45, "CFO\nTGO\nEF ", "1", "L", false)

	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(1)  ข้อมูลกิจกรรมที ่ได้จากการตรวจวัด ให้ระบุรายละเอียดการสอบเทียบของอุปกรณ์ตรวจวัดไว้ในตารางที ่ 7.3   ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(2) ข้อมูลกิจกรรมที่ได้จากการประมาณค่า ให้อธิบายแนวทางในการประมาณในตารางหรืออธิบายเพิ่มเติมในภาคผนวก  ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(3) ในกรณีที ่ข้อมูลกิจกรรมเป็นข้อมูลปริมาณการปล่อยก๊าซเรือนกระจกอยู่แล้ว เช่น ปริมาณการรั ่วซึมของสารท าความเย็น ให้กรอกค าว่า “ไม่ต้องใช้ค่า EF” ลงในคอลัมน์ “ที ่มาของค่า EF” ", "", 2, "L", false, 0, "")

	// 4.4
	// Set text color to black

	pdf.SetTextColor(0, 0, 0)
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, "4.4 แหล่งปล่อยก๊าซเรือนกระจก จากขอบเขตการด าเนินงานประเภทรายงานแยกเพิ ่มเติม", "", 1, "L", false, 0, "")

	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "ในกรณีที ่รำยงำนก๊ำซเรื ่อนกระจกอื ่น ๆเพิ ่มเติม หรือ รำยงำนแยกในส่วนของไบโอจินิคคำร์บอน (ถ้ำมี) ", "", 2, "L", false, 0, "")

	pdf.SetTextColor(0, 0, 0)
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(35, 7.5, "", "TL", 0, "L", true, 0, "")
	pdf.CellFormat(110, 7.5, "ข้อมูลกิจกรรม", "1", 0, "C", true, 0, "")
	pdf.SetFont("THSarabunNew", "B", 12)
	pdf.MultiCell(10, 7.5, "ค่า EF", "1", "L", true)

	pdf.SetFont("THSarabunNew", "", 12)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 22.5, "แหล่งปล่อยก๊าซเรือนกระจก ", "BL", "C", true)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 7.5, "ลักษณะข้อมูลกิจกรรมที่ตรวจวัด ", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 11.25, "จุดที่\nตรวจวัด", "1", "L", true)
	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(54, 7.5, "ที่มาของข้อมูลกิจกรรม ", "1", "C", true)
	pdf.SetXY(x, y+7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 7.5, "เป็นค่าที่ได้จากการตรวจวัด", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากหลักฐานการชำระเงิน", "1", "L", true)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 5, "เป็นค่าที่ได้จากการประเมินค่า", "1", "L", true)
	pdf.SetXY(x+18, y-7.5)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 11.25, "หลักฐาน/\nเอกสารอ้างอิง", "1", "L", true)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 7.5, "ที่มา\nของค่า EF", "1", "L", true)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 7.5, "-ไม่มี- ", "BL", "C", false)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 7.5, "", "1", "C", false)

	pdf.SetXY(x+15, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(15, 7.5, "", "1", "L", false)
	pdf.SetXY(x+15, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(18, 7.5, "", "1", "L", false)

	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 7.5, "", "1", "L", false)

	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(18, 7.5, "", "1", "L", false)
	pdf.SetXY(x+18, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(26, 7.5, "", "1", "L", false)
	pdf.SetXY(x+26, y)

	pdf.MultiCell(10, 7.5, "", "1", "L", false)

	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "หมายเหตุ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(1)  ข้อมูลกิจกรรมที ่ได้จากการตรวจวัด ให้ระบุรายละเอียดการสอบเทียบของอุปกรณ์ตรวจวัดไว้ในตารางที ่ 7.3   ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(2) ข้อมูลกิจกรรมที่ได้จากการประมาณค่า ให้อธิบายแนวทางในการประมาณในตารางหรืออธิบายเพิ่มเติมในภาคผนวก  ", "", 2, "L", false, 0, "")
	pdf.CellFormat(45, 7.5, "(3) ในกรณีที ่ข้อมูลกิจกรรมเป็นข้อมูลปริมาณการปล่อยก๊าซเรือนกระจกอยู่แล้ว เช่น ปริมาณการรั ่วซึมของสารท าความเย็น ให้กรอกค าว่า “ไม่ต้องใช้ค่า EF” ลงในคอลัมน์ “ที ่มาของค่า EF” ", "", 2, "L", false, 0, "")

	//5.1
	pdf.SetTextColor(0, 0, 0)
	pdf.SetFont("THSarabunNew", "B", 16)
	//pdf.AddPage()
	pdf.CellFormat(0, 10, "5. สรุปปริมาณการปล่อยก๊าซเรือนกระจก", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, " 5.1 การปล่อยก๊าซเรือนกระจก จากขอบเขตการด าเนินงานประเภทที ่ 1 ", "", 1, "L", false, 0, "")

	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "เฉพำะประเภทที ่ 1 ให้แยกชนิดก๊ำซในแต่ละแหล่งปล่อย", "", 2, "L", true, 0, "")

	pdf.SetTextColor(0, 0, 0)
	pdf.SetFont("THSarabunNew", "B", 10)
	previousX, previousY := pdf.GetXY()
	pdf.CellFormat(35, 21.75, "แหล่งปล่อยก๊าซเรือนกระจก", "TL", 0, "L", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(110, 7, "ปริมาณการปล่อยก๊าซเรือนกระจก \n(Ton CO2e)", "1", "C", true)
	pdf.SetXY(x+110, y)
	pdf.MultiCell(30, 7.25, "รวมปริมาณ\nก๊าซเรือนกระจก \n(Ton CO2e) ", "1", "L", true)

	pdf.SetFont("THSarabunNew", "", 10)
	//pdf.Ln(-1)

	pdf.SetXY(previousX+35, previousY+14)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "CO2 ", "1", "C", true)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "Fossil CH4", "1", "C", true)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "CH4 ", "1", "C", true)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "N2O", "1", "C", true)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "SF6", "1", "C", true)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "NF3", "1", "C", true)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "HFCs ", "1", "C", true)
	pdf.SetXY(x+13.75, y)

	pdf.MultiCell(13.75, 7.5, "PFCs ", "1", "C", true)

	x, y = pdf.GetXY()
	pdf.MultiCell(35, 7.5, "Mobile Combustion ", "BL", "C", false)
	pdf.SetXY(x+35, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "C", false)

	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "L", false)
	pdf.SetXY(x+13.75, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "L", false)

	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "L", false)

	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "L", false)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "L", false)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "L", false)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "", "1", "L", false)
	pdf.SetXY(x+13.75, y)

	pdf.MultiCell(30, 7.5, "", "1", "L", false)

	x, y = pdf.GetXY()
	pdf.MultiCell(5, 7.5, "1", "BL", "C", false)
	pdf.SetXY(x+5, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(30, 7.5, "น ้ามันดีเซลรถยนต์ ", "BL", "L", false)
	pdf.SetXY(x+30, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "316.86 ", "1", "C", false)

	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "0.53 ", "1", "L", false)
	pdf.SetXY(x+13.75, y)

	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "-", "1", "L", false)

	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "31.45", "1", "L", false)

	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "-", "1", "L", false)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "-", "1", "L", false)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "-", "1", "L", false)
	pdf.SetXY(x+13.75, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(13.75, 7.5, "-", "1", "L", false)
	pdf.SetXY(x+13.75, y)

	pdf.MultiCell(30, 7.5, "348.84 ", "1", "L", false)

	//5.2
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, "5.2 การปล่อยก๊าซเรือนกระจก จากขอบเขตการด าเนินงานประเภทที ่ 2 ", "", 1, "L", false, 0, "")

	// 5.2 table
	//pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 14)

	data = [][]string{
		{"แหล่งปล่อยก๊าซเรือนกระจก ", "ปริมาณการปล่อย GHG (Ton CO2e)"},
		{"การใช้ไฟฟ้า", "28,252.52 "},
		{"รวมทั ้งหมด ", "28,253 "},
	}
	//generateTableContent(pdf, data, []float64{60.0, 120.0})

	//5.3
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, "5.3 การปล่อยก๊าซเรือนกระจก จากขอบเขตการด าเนินงานประเภทที ่ 3", "", 1, "L", false, 0, "")

	// 5.3 table
	//pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 14)

	data = [][]string{
		{"แหล่งปล่อยก๊าซเรือนกระจก ", "ปริมาณการปล่อย GHG (Ton CO2e)"},
		{"1. Purchased goods and services ", "1,269,288.86"},
		{"2. Fuel and energy related activities not included scope 1 & 2", "8,018.48"},
		{"3. Upstream transport ", "83,586.95 "},
		{"4. Waste generate ", "263.95"},
		{"5. Downstream transport ", "19,068.44"},
		{"รวมทั ้งหมด ", "1,380,227"},
	}
	//generateTableContent(pdf, data, []float64{60.0, 120.0})

	// 5.4 table
	pdf.AddPage()
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.MultiCell(0, 10, "5.4 การปล่อยก๊าซเรือนกระจก จากขอบเขตการด าเนินงานประเภทที ่รายงานแยกเพิ ่มเติม ", "", "L", false)
	// Set text color to blue
	pdf.SetTextColor(0, 0, 255)
	pdf.SetFont("THSarabunNew", "", 10)
	pdf.CellFormat(45, 7.5, "ในกรณีที ่รำยงำนก๊ำซเรื ่อนกระจกอื ่น ๆเพิ ่มเติม หรือ รำยงำนแยกในส่วนของไบโอจินิคคำร์บอน (ถ้ำมี)", "", 2, "L", true, 0, "")

	pdf.SetTextColor(0, 0, 0)

	data = [][]string{
		{"อุปกรณ์ / เครื่องจักรที่ผลิตพลังงาน / ความร้อน / ไอน้ำ / กระบวนการ (Source)  ", "จำหน่ายให้กับ (Supply to) "},
		{"-ไม่มี-", " "},
	}
	//generateTableContent(pdf, data, []float64{60.0, 120.0})

	//5.5 table
	pdf.MultiCell(0, 10, "5.5 Carbon Intensity ", "", "L", false)
	pdf.SetFillColor(190, 190, 190)
	pdf.CellFormat(50, 15, "แหล่งปล่อยก๊าซเรือนกระจก  ", "1", 0, "C", true, 0, "")
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 15, "ปริมาณ", "1", "C", true)
	pdf.SetXY(x+50, y)

	pdf.MultiCell(50, 15, "หน่วย ", "1", "C", true)

	x, y = pdf.GetXY()
	pdf.MultiCell(50, 6, "ประเภทที่ 1", "1", "C", false)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 6, "45,065.00", "1", "C", false)
	pdf.SetXY(x+50, y)

	pdf.MultiCell(50, 6, "Ton CO2e", "1", "C", false)

	//6 / 6.1

	pdf.SetTextColor(0, 0, 0)
	pdf.SetFont("THSarabunNew", "B", 16)
	//pdf.AddPage()
	pdf.CellFormat(0, 10, "6. ปีฐาน ", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, " 6.1 ปีฐานที ่ใช้ในการอ้างอิง ", "", 1, "L", false, 0, "")
	generateTextContent(pdf, true, "มกราคม ถึง ธันวาคม 2564 ซึ ่งเป็นข้อมูลที ่ได้รับการทวนสอบความถูกต้องจากผู ้ทวนสอบเรียบร้อยแล้ว โดยคลอบคลุมพื ้นที ่ รายละเอียดตามรายงานข้อ 3.1.4 ของรายงานฉบับนี ้")

	//6.2
	pdf.AddPage()
	pdf.CellFormat(0, 10, "6.2 ขอบเขตการด าเนินงานในปีฐาน", "", 1, "L", false, 0, "")

	//6.2 table
	x, y = pdf.GetXY()
	pdf.MultiCell(25, 7.5, "ขอบเขต\nการดำเนินงาน   ", "1", "C", true)
	pdf.SetXY(x+25, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 7.5, "รายการแหล่งปล่อย\nก๊าซเรือนกระจก ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 7.5, "ปริมาณการปล่อยก๊าซเรือนกระจก\nของปีฐาน(Ton CO2e)    ", "1", "C", true)
	pdf.SetXY(x+50, y)
	pdf.MultiCell(30, 15, "หมายเหตุ", "1", "C", true)

	x, y = pdf.GetXY()
	pdf.MultiCell(25, 6*16, "ขอบเขตที่ 1", "1", "C", false)
	pdf.SetXY(x+25, y)

	for i := range 16 {
		//log.Println(i)
		x, y = pdf.GetXY()
		pdf.MultiCell(50, 6, strconv.Itoa(i+1)+". น้ำมันดีเซลรถยนต์", "1", "C", false)
		pdf.SetXY(x+50, y)
		pdf.MultiCell(50, 6, "298.64 ", "1", "C", false)
		pdf.SetXY(x+100, y)
		pdf.MultiCell(30, 6, "", "1", "C", false)
		pdf.SetXY(x, y+6)
	}

	//7.
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, "7. การจัดการคุณภาพของข้อมูล", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, " 7.1 โครงสร้างของระบบการจัดการคุณภาพของข้อมูล ", "", 1, "L", false, 0, "")

	//7.1 table
	pdf.SetFont("THSarabunNew", "", 14)
	x, y = pdf.GetXY()
	pdf.MultiCell(25, 7.5, "บทบาท ", "1", "C", true)
	pdf.SetXY(x+25, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 7.5, "ชื่อ-สกุล ", "1", "C", true)
	pdf.SetXY(x+50, y)
	x, y = pdf.GetXY()
	pdf.MultiCell(50, 7.5, "ตำแหน่ง  ", "1", "C", true)
	pdf.SetXY(x+50, y)
	pdf.MultiCell(40, 7.5, "หน้าที่ ", "1", "C", true)

	data = [][]string{
		{"คุณ ไกรศกด กลบทอง", "ผู้จัดการโรงงาน BTG LR1"}, {"คุณ ภาคภูมิ สีแก้วสิ่ว ", "ผู้จัดการฝ่ายผลิตโรงงาน "},
		{"คุณ บรรจบ ศฤงคารินทร์ ", "ผู้จัดการโรงงาน BTG LR3"},
	}

	firstColX, firstColY := pdf.GetXY()

	pdf.SetX(pdf.GetX() + 25)

	for _, name := range data {
		//log.Println(i)
		x, y = pdf.GetXY()
		pdf.MultiCell(50, 6, name[0], "1", "C", false)
		pdf.SetXY(x+50, y)
		pdf.MultiCell(50, 6, name[1], "1", "C", false)
		pdf.SetXY(x+100, y)
		pdf.MultiCell(30, 6, "", "", "C", false)
		pdf.SetXY(x, y+6)

	}
	pdf.SetXY(x+100, firstColY)
	pdf.MultiCell(40, (6 * float64(len(data)) / 2), "กำหนดนโยบายในการ\nบริหารงานขององค์กร ", "1", "C", false)

	pdf.SetXY(firstColX, firstColY)
	pdf.MultiCell(25, (6 * float64(len(data)) / 2), "ผู้จัดการ\nโรงงาน ", "1", "C", false)
	pdf.SetXY(x+25, y)

	//7.2
	//1.
	//1.1
	pdf.SetFont("THSarabunNew", "B", 16)

	pdf.CellFormat(0, 10, "7.2 แผนผังการจัดการคุณภาพของข้อมูล ", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, "Scope 1 ", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, "1. Mobile Combustion", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, " 1.1 น้ำมันดีเซลรถยนต์ ", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"ใบรายงานการใช้น้ำมัน\nและแก๊ส/ใบเสร็จรับเงิน ", "เจ้าหน้าที่\nแผนกทรัพยากรมนุษย์ \nความถี่ : เดือนละ 1 ครั้ง ", "ผู้จัดการ\nแผนกทรัพยากรมนุษย์\nความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่\nสิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
		{"ยอดเบิกจาก SAP (เบิกจาก Station ภายในโรงงาน) ", "พนักงานจ่ายน้ำมัน \nความถี่ : ทุกครั้งที่มีการเติม ", "เจ้าหน้าที่สโตร์\nความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
		{"Fleet Card", "เจ้าหน้าที่แผนกยานยนต์ \nส่วนกลาง เบทาโกร \nความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่แผนกบัญชี\nส่วนกลาง เบทาโกร \nความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//1.2
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, " 1.2 น้ำมันเบนซีนรถยนต์ ", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"รายงานสรุป Fleet Card ", "เจ้าหน้าที่แผนกยานยนต์ \nความถี่ : เดือนละ 1 ครั้ง ", "ผู้จัดการแผนกยานยนต์ \nความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
		{"1.ใบรายงานการใช้น้ำมัน\nและแก๊ส/ใบเสร็จรับเงิน\n2.ราคาน้ำมันเฉลี่ยรายเดือน\n3.SAP", "เจ้าหน้าที่แผนกทรัพยากรมนุษย์ ความถี่ : เดือนละ 1 ครั้ง", "ผู้จัดการแผนกทรัพยากรมนุษย์ ความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//1.3
	pdf.SetFont("THSarabunNew", "B", 16)

	pdf.CellFormat(0, 10, " 1.3 ก๊าซ NGV สำหรับรถยนต์ ", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"1.SAP\n2.ราคา NVG เฉลี่ยต่อเดือน", "เจ้าหน้าที่แผนกทรัพยากรมนุษย์ ความถี่ : เดือนละ 1 ครั้ง", "ผู้จัดการแผนกทรัพยากรมนุษย์ ความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//1.4
	pdf.SetFont("THSarabunNew", "B", 16)

	pdf.CellFormat(0, 10, " 1.4 ก๊าซ LPG รถโฟร์คลิฟ ", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"ยอดเบิกจากระบบ SAP", "ผู้จัดการแผนกคลังสินค้า ความถี่ : ทุกครั้งที่มีการเบิก", " ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//2.
	//2.1
	pdf.SetFont("THSarabunNew", "B", 16)

	pdf.CellFormat(0, 10, "2. Stationary Combustion ", "", 1, "L", false, 0, "")
	pdf.CellFormat(0, 10, " 2.1 น้ำมันดีเซล Fire pump ", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"ยอดเบิกจาก SAP (เบิกจาก Station ภายในโรงงาน) ", "พนักงานจ่ายน้ำมัน \nความถี่ : ทุกครั้งที่มีการเติม ", "เจ้าหน้าที่สโตร์\nความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//2.2
	pdf.SetFont("THSarabunNew", "B", 16)
	// pdf.AddPage()
	pdf.CellFormat(0, 10, " 2.2 น้ำมันเบนซีนเครื่องตัดหญ้า ", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"SAP", "เจ้าหน้าที่แผนกทรัพยากรมนุษย์ ความถี่ : เดือนละ 1 ครั้ง", "ผู้จัดการแผนกทรัพยากรมนุษย์ ความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//2.3
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, " 2.3 น้ำมันเตา C", "", 1, "L", false, 0, "")

	pdf.SetFont("THSarabunNew", "", 14)
	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"ยอดเบิกจาก SAP (Boiler)", "เจ้าหน้าที่ซ่อมบำรุง (ดูแล Boiler) \nความถี่ : เดือนละ 1 ครั้ง", "เจ้าหน้าที่ธุรการผลิต \nความถี่ : เดือนละ 1 ครั้ง ", "เจ้าหน้าที่สิ่งแวดล้อม \nความถี่ : เดือนละ 1 ครั้ง"},
		{"ยอดเบิกจาก SAP (อบข้าวโพด)", "เจ้าหน้าที่ silo ความถี่ : ทุกครั้งที่มีการเบิก", " ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//2.4
	pdf.SetFont("THSarabunNew", "B", 16)

	pdf.CellFormat(0, 10, " 2.4 ก๊าซ LPG ซ่อมบำรุง", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"ยอดเบิกจากระบบ SAP LR1,LR2,LR3 ", "ผู้จัดการแผนกคลังสินค้า ความถี่ : ทุกครั้งที่มีการเบิก", " ", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//2.5
	pdf.SetFont("THSarabunNew", "B", 16)

	pdf.CellFormat(0, 10, " 2.5 ถ่านหิน Boiler", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"1.ยอดเบิกจากระบบ SAP\n2.ค่าความร้อนจาก Supplier", "เจ้าหน้าที่สโตร์ ความถี่ : เดือนละ 1 ครั้ง", "ผู้จัดการผลิต \n ความถี่ : เดือนละ 1 ครั้ง", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	//generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	//7.2
	//2.6
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.AddPage()
	pdf.CellFormat(0, 10, " 2.6 ก๊าซ LPG for Boiler", "", 1, "L", false, 0, "")

	data = [][]string{
		{"หลักฐานอ้างอิง ", "การบันทึกข้อมูล", "การตรวจสอบข้อมูล", "การรวบรวมข้อมูลคำนวณ CFO"},
		{"ยอดเบิกใช้จากระบบ SAP LR3", "เจ้าหน้าที่สโตร์ ความถี่ : ทุกครั้งที่มีการเบิก", "ผู้จัดการผลิต \n ความถี่ : เดือนละ 1 ครั้ง", "เจ้าหน้าที่สิ่งแวดล้อม\nความถี่ : เดือนละ 1 ครั้ง"},
	}
	generateTableContent(pdf, data, []float64{40.0, 40.0, 40.0, 40.0})

	// ภาคผนวก
	pdf.AddPage()
	// Title
	pdf.SetFont("THSarabunNew", "B", 16)
	pdf.CellFormat(0, 10, "ภาคผนวก", "", 1, "C", false, 0, "")

	// Content
	// paragraph 1
	pdf.SetFont("THSarabunNew", "B", 14)
	content = `ตามข้อก าหนดของ อบก. ก าหนดให้องค์กรมีกระบวนการชี ้บ่งแหล่งปล่อยก๊าซเรือนกระจกทางอ้อมอื ่นๆ (ประเภทที ่ 3) ที ่จะน ามารวมในบัญชีรายการก๊าซเรือนกระจก โดยให้ความส าคัญของแหล่งการปล่อยก๊าซเรือนกระจกตามหลักเกณฑ์ดังต่อไปนี ้ `
	generateTextContent(pdf, true, content)

	// paragraph 2
	content = `- ขนาด (Magnitude): เป็นกิจกรรมการปล่อยหรือดูดกลับก๊าซเรือนกระจกทางอ้อมซึ่่งถูกสันนิษฐานว่ามีปริมาณการปล่อยหรือดูดกลับก๊าซเรือนกระจกในปริมาณมากอย่างมีนัยส าคัญ `
	generateTextContent(pdf, true, content)

	// paragraph 3
	content = `- ระดับของแรงจูงใจ(Level of influence): เป็นกิจกรรมการปล่อยหรือดูดกลับก๊าซเรือนกระจกที ่องค์กรมีความสามารถในการตรวจติดตามและลดปริมาณการปล่อยหรือดูดกลับก๊าซเรือนกระจกจากกิจกรรมนั ้น(ตัวอย่างเช่นเป็นกิจจกรรมที ่เกี ่ยวข้องกับการประเมินประสิทธิภาพพลังงาน การออกแบบ ชิงนิเวศเศรษฐกิจ, เกี ่ยวข้องกับข้อตกลงที ่มีกับลูกค้า, เกี ่ยวข้องกับข้อก าหนดขอบเขตงานจากผู ้ว่าจ้าง)`
	generateTextContent(pdf, true, content)

	// paragraph 4
	content = `- ความเสี ่ยงหรือโอกาส (Risk or opportunity): เป็นกิจกรรมการปล่อยหรือดูดกลับก๊าซเรือนกระจกทางอ้อมซึ ่งมีส่วนท าให้องค์กรได้รับความเสี ่ยง (ตัวอย่างของความเสี ่ยงที ่มีความเชื ่อมโยงกับการเปลี ่ยนแปลงสภาพภูมิอากาศ เช่น ความเสี ่ยงทางด้านการเงิน, ความเสี ่ยงทางด้านกฎระเบียบข้อบังคับ, 
ความเสี ่ยงตลอดห่วงโซ่อุปทาน, ความเสี ่ยงเกี ่ยวกับสินค้าและลูกค้า, ความเสี ่ยงเกี ่ยวกับการด าเนินคดี และ ความเสี ่ยงด้านชื ่อเสียง) หรือได้รับโอกาสต่างๆ ทางธุรกิจ (เช่น การเข้าสู ่ช่องทางตลาดใหม่ การเข้าสู ่ระบบธุรกิจในรูปแบบใหม่) `
	generateTextContent(pdf, true, content)

	// paragraph 5
	content = `- เป็นการจัดจ้างบุคคลหรือหน่วยงานภายนอก (Outsourcing): เป็นกิจกรรมการปล่อยและดูดกลับก๊าซเรือนกระจกทางอ้อมที ่เกิดจากการจัดจ้างบุคคลหรือหน่วยงานภายนอกเข้ามาด าเนินกิจกรรมที ่ถือว่าเป็นกิจกรรมหลักในการด าเนินธุรกิจขององค์กร  `
	generateTextContent(pdf, true, content)

	// paragraph 6
	content = `- เป็นการส่งเสริมการมีส่วนร่วมของพนักงาน (Employee engagement): เป็นกิจกรรมการปล่อยก๊าซเรือนกระจกทางอ้อมที ่สามารถส่งเสริมให้เกิดการกระตุ ้นให้พนักงานมีส่วนร่วมในการลดการปล่อยก๊าซเรือนกระจก ผ่านการลดการใช้พลังงาน หรือการท างานร่วมกันเป็นทีมภายใต้หลักคิดที ่เกี ่ยวข้องกับการเปลี ่ยนแปลงสภาพภูมิอากาศ (เช่น การสร้างแรงจูงใจในการอนุรักษ์พลังงาน, การเดินทางโดยใช้รถร่วมกัน, การประเมินราคาคาร์บอนภายในองค์กร เป็นต้น)  `
	generateTextContent(pdf, true, content)

	// Save the PDF to a file
	err := pdf.OutputFileAndClose("output.pdf")
	if err != nil {
		fmt.Println("Error saving PDF:", err)
		return
	}

	fmt.Println("PDF created successfully")
}
