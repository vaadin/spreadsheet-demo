package com.vaadin.addon.spreadsheet.demo.examples;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.SpreadsheetComponentFactory;
import com.vaadin.addon.spreadsheet.action.SpreadsheetDefaultActionHandler;
import com.vaadin.event.Action.Handler;
import com.vaadin.ui.Component;

public class ComponentsExample implements SpreadsheetExample {

    private Spreadsheet spreadsheet;
    private SpreadsheetComponentFactory spreadsheetFieldFactory = new TestComponentFactory(
            this);
    private Handler spreadsheetActionHandler = new SpreadsheetDefaultActionHandler();

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
        spreadsheet.addActionHandler(spreadsheetActionHandler);
    }

    public Spreadsheet getSpreadsheet() {
        return spreadsheet;
    }

}
