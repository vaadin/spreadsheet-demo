package com.vaadin.addon.spreadsheet.demo.helpers;

import java.io.File;
import java.io.IOException;

import com.vaadin.addon.spreadsheet.Spreadsheet;

public class FileExampleHelper {

    // opens the specified file as a spreadsheet
    public static Spreadsheet openFile(File file) {
        Spreadsheet spreadsheet = null;
        try {
            spreadsheet = new Spreadsheet(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return spreadsheet;
    }

}
