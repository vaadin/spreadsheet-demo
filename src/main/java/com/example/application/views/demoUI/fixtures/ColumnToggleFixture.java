package com.example.application.views.demoUI.fixtures;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.CellReference;

import com.vaadin.flow.component.spreadsheet.Spreadsheet;


public class ColumnToggleFixture implements SpreadsheetFixture {

    @Override
    public void loadFixture(Spreadsheet spreadsheet) {
        List<Integer> columnIndexes = new ArrayList<Integer>();
        for (CellReference cellRef : spreadsheet.getSelectedCellReferences()) {
            if (!columnIndexes.contains((int) cellRef.getCol())) {
                columnIndexes.add((int) cellRef.getCol());
            }
        }

        for (Integer col : columnIndexes) {
            spreadsheet.setColumnHidden(col, !spreadsheet.isColumnHidden(col));
        }

        spreadsheet.refreshAllCellValues();
    }
}
