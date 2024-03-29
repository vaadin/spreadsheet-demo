package com.example.application.views.demoUI.fixtures;

import org.apache.poi.ss.usermodel.Workbook;

import com.vaadin.flow.component.spreadsheet.Spreadsheet;


public class RenameFixture implements SpreadsheetFixture {
    @Override
    public void loadFixture(Spreadsheet spreadsheet) {
        spreadsheet.getCell(0, 0);
        Workbook wb = spreadsheet.createCell(0, 0, "").getSheet()
                .getWorkbook();
        wb.getSheet("new sheet name");
        spreadsheet.setSheetName(wb.getSheetIndex("new sheet name"),
                "new_sheet_REnamed");
    }
}
