package main

import (
	"fmt"
	_ "image/png"

	"github.com/xuri/excelize/v2"
)

func fillTemplate(template string) {
	f, err := excelize.OpenFile(fmt.Sprintf("template_%s.xlsx", template))
	if err != nil {
		fmt.Println("Could not open template", err)
		return
	}
	if err := f.AddPicture("INDEX", "A1", "image.png", &excelize.GraphicOptions{OffsetX: 10, OffsetY: 10, ScaleX: 0.3, ScaleY: 0.3}); err != nil {
		fmt.Println("Could add picture to INDEX", err)
		return
	}
	var templateIndex = 0
	for index, sh := range f.WorkBook.Sheets.Sheet {
		if sh.Name == "TEMPLATE" {
			templateIndex = index
		}
	}
	if templateIndex == 0 {
		fmt.Println("TEMPLATE sheet not found", err)
		return
	}
	itemIndex, err := f.NewSheet("Item1")
	if err != nil {
		fmt.Println("Could not create sheet", err)
		return
	}
	f.CopySheet(templateIndex, itemIndex)

	if err := f.AddPicture("Item1", "A1", "image.png", &excelize.GraphicOptions{OffsetX: 10, OffsetY: 10, ScaleX: 0.3, ScaleY: 0.3}); err != nil {
		fmt.Println("Could add picture to Item1", err)
		return
	}
	f.DeleteSheet("TEMPLATE")
	err = f.SaveAs(fmt.Sprintf("./out/%s.xlsx", template))
	if err != nil {
		fmt.Println("Could not save file", template, err)
		return
	}
}

func main() {
	fillTemplate("simple")
	fillTemplate("broken")
	fillTemplate("fixed")
}
