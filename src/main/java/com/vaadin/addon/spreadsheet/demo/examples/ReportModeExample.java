package com.vaadin.addon.spreadsheet.demo.examples;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.demo.SpreadsheetDemoUI;
import com.vaadin.ui.Component;

public class ReportModeExample implements SpreadsheetExample {

    private Spreadsheet spreadsheet;

    public ReportModeExample() {

        initSpreadsheet();
    }

    @Override
    public Component getComponent() {
        return spreadsheet;
    }

    private void initSpreadsheet() {
        File sampleFile = null;
        try {
            ClassLoader classLoader = SpreadsheetDemoUI.class.getClassLoader();
            URL resource = classLoader.getResource(
                    "testsheets" + File.separator + "Simple Invoice.xlsx");
            if (resource != null) {
                sampleFile = new File(resource.toURI());
                spreadsheet = new Spreadsheet(sampleFile);
                spreadsheet.setReportStyle(true);
                spreadsheet.setActiveSheetProtected("");
                spreadsheet.setRowColHeadingsVisible(false);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
    }

}
