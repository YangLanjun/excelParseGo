package main

import (
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"os"
	"strconv"
)

type person struct {
	 name string
	 address string
	 cardId string
	 money float64

}

var personList []person=make([]person,0)

var id int=9
var index int=1
var rid int=7

var images=make([]string,0)

func main() {
	parseExcel("C:\\Users\\10691\\Desktop\\1.xlsx")
	writeWord()
}
func rowVisitor(r *xlsx.Row) error {
	var p person
	p.name=r.GetCell(2).Value
	p.address=r.GetCell(3).Value
	p.cardId = r.GetCell(4).Value
	var err error
	p.money,err=r.GetCell(5).Float()
	if err!=nil{
		p.money=personList[len(personList)-1].money
	}else{
		p.money=Decimal(p.money/10000)
	}
	personList = append(personList, p)
	return nil
}

func parseExcel(filePath string) {
	excelFile, error := xlsx.OpenFile(filePath)
	if error != nil {
		fmt.Errorf(error.Error())
		return
	}
	sheet := excelFile.Sheets[0]
	err := sheet.ForEachRow(rowVisitor)
	printError(err)
	for _,p:=range personList{
		printPerson(p)
	}
}

func printError(err error){
	if err!=nil{
		fmt.Println("Err=", err)
	}
}

func printPerson(p person){
	fmt.Printf("%s,%s,%s,%f\n",p.name,p.address,p.cardId,p.money)
}

func Decimal(value float64) float64 {
	value, _ = strconv.ParseFloat(fmt.Sprintf("%.2f", value), 64)
	return value
}


func writeWord() {
	fname:="document.xml"
	os.Remove(fname)
	file, _ := os.OpenFile(fname, os.O_RDWR | os.O_CREATE, 0664)
	defer file.Close()
	sHead:="<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:wpsCustomData=\"http://www.wps.cn/officeDocument/2013/wpsCustomData\" mc:Ignorable=\"w14 w15 wp14\"><w:body><w:tbl><w:tblPr><w:tblStyle w:val=\"9\"/><w:tblW w:w=\"8816\" w:type=\"dxa\"/><w:tblInd w:w=\"-176\" w:type=\"dxa\"/><w:tblBorders><w:top w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:left w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:bottom w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:right w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideH w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideV w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/></w:tblBorders><w:tblLayout w:type=\"fixed\"/><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w=\"4378\"/><w:gridCol w:w=\"4438\"/></w:tblGrid>"
	file.WriteString(sHead)
	for _,p:=range personList{
		str:=buildTable(p)
		file.WriteString(str)
	}
	sFoot:="</w:tbl><w:p><w:pPr><w:pStyle w:val=\"6\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFFFF\"/><w:spacing w:line=\"240\" w:lineRule=\"auto\"/><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:pPr></w:p><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\" w:header=\"851\" w:footer=\"992\" w:gutter=\"0\"/><w:cols w:space=\"425\" w:num=\"1\"/><w:docGrid w:type=\"lines\" w:linePitch=\"312\" w:charSpace=\"0\"/></w:sectPr></w:body></w:document>"
	file.WriteString(sFoot)
	file.Sync()

	refName:="document.xml.rels"
	os.Remove(refName)
	refFile, _ := os.OpenFile(refName, os.O_RDWR | os.O_CREATE, 0664)
	defer refFile.Close()
	rsHead:="<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
	refFile.WriteString(rsHead)
	for i:=len(images)-1;i>=0;i--{
		refFile.WriteString(images[i])
	}
	rsFoot:="<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item2.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item1.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>"
	refFile.WriteString(rsFoot)

}

func buildTable(p person) string{
	imageName:=p.cardId+p.name
	s1:="<w:tr><w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:left w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:bottom w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:right w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideH w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideV w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/></w:tblBorders><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPrEx><w:trPr><w:trHeight w:val=\"4351\" w:hRule=\"atLeast\"/></w:trPr><w:tc><w:tcPr><w:tcW w:w=\"4378\" w:type=\"dxa\"/></w:tcPr><w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/>" +
		"</w:rPr><w:t>"+strconv.Itoa(index)+"</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"1540510\" cy=\"1896110\"/><wp:effectExtent l=\"19050\" t=\"0\" r=\"2241\" b=\"0\"/>" +
		"<wp:docPr id=\""+strconv.Itoa(id)+"\" name=\"图片 "+strconv.Itoa(index)+"\" descr=\""+imageName+"\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
		"<pic:nvPicPr><pic:cNvPr id=\""+strconv.Itoa(id)+"\" name=\"图片 "+strconv.Itoa(index)+"\" descr=\""+imageName+
		"\"/><pic:cNvPicPr><a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>" +
		"</pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId"+strconv.Itoa(rid)+"\" cstate=\"print\"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1544522\" cy=\"1901193\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"9525\"><a:noFill/><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:tc><w:tc><w:tcPr>" +
		"<w:tcW w:w=\"4438\" w:type=\"dxa\"/></w:tcPr><w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/>" +
		"<w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/>" +
		"<w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr><w:t>失信</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/>" +
		"</w:rPr><w:t>被执行</w:t></w:r><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"0\"/><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/>" +
		"<w:szCs w:val=\"32\"/></w:rPr><w:t>人：</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/>" +
		"</w:rPr><w:t>"+p.name+"</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr><w:t xml:space=\"preserve\"> " +
		"</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:pPr><w:r><w:rPr>" +
		"<w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr><w:t>公民</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/>" +
		"</w:rPr><w:t>身份证号码："+p.cardId[:6]+"********"+p.cardId[14:]+"</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/>" +
		"<w:szCs w:val=\"32\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/>" +
		"<w:szCs w:val=\"32\"/></w:rPr><w:t>住址</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/>" +
		"<w:szCs w:val=\"32\"/></w:rPr><w:t>：</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/>" +
		"</w:rPr><w:t>"+p.address+"</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr>" +
		"<w:t></w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr><w:t>应</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr><w:t>履行法定义务：</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:asciiTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorEastAsia\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr>" +
		"<w:t>"+ strconv.FormatFloat(p.money, 'f', -1, 64)+"万元</w:t></w:r></w:p></w:tc></w:tr>"

	buildImage(p)
	sImage:="<Relationship Id=\"rId"+strconv.Itoa(rid)+"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image"+strconv.Itoa(index)+".jpeg\"/>"
	images = append(images, sImage)
	rid++
	index++
	id++
	return s1
}

func buildImage(p person)int{
	fileName:="../1/word/media/"+p.cardId+p.name+".jpg"
	println(fileName)
	err:=os.Rename(fileName,"../1/word/media/"+"image"+strconv.Itoa(index)+".jpeg")
	if err!=nil{
		fmt.Println("rename successfully")
	}
	return index
}
