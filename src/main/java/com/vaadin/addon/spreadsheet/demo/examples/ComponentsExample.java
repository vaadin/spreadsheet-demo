package com.vaadin.addon.spreadsheet.demo.examples;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.SpreadsheetComponentFactory;
import com.vaadin.ui.Component;

public class ComponentsExample implements SpreadsheetExample {

    private Spreadsheet spreadsheet;
    private SpreadsheetComponentFactory spreadsheetFieldFactory = new TestComponentFactory(
            this);

    public ComponentsExample() {

        initSpreadsheet();
    }

    @Override
    public Component getComponent() {
        return spreadsheet;
    }

    private void initSpreadsheet() {
        spreadsheet = new Spreadsheet(
                ((TestComponentFactory) spreadsheetFieldFactory)
                        .getTestWorkbook());
        spreadsheet.setSpreadsheetComponentFactory(spreadsheetFieldFactory);
    }

    public Spreadsheet getSpreadsheet() {
        return spreadsheet;
    }

}
