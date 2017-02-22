package com.vaadin.addon.spreadsheet.demo.examples;

import java.text.Format;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.ExcelStyleDateFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.SpreadsheetComponentFactory;
import com.vaadin.shared.ui.ContentMode;
import com.vaadin.shared.ui.datefield.DateResolution;
import com.vaadin.ui.Button;
import com.vaadin.ui.Button.ClickEvent;
import com.vaadin.ui.CheckBox;
import com.vaadin.ui.ComboBox;
import com.vaadin.ui.Component;
import com.vaadin.ui.DateField;
import com.vaadin.ui.Label;
import com.vaadin.ui.NativeSelect;
import com.vaadin.ui.Notification;

@SuppressWarnings("serial")
public class TestComponentFactory implements SpreadsheetComponentFactory {

    private int counter = 0;

    private final DateField dateField = new DateField();

    private final CheckBox  checkBox = new CheckBox();

    private final Workbook testWorkbook;

    private final String[] comboBoxValues = { "Value 1", "Value 2", "Value 3" };

    private final Object[][] data = {
            { "Testing custom editors", "Boolean", "Date", "Numeric", "Button",
                    "ComboBox" },
            { "nulls:", false, null, 0, null, null },
            { "", true, new Date(), 5, "here is a button", comboBoxValues[0] },
            { "", true, Calendar.getInstance(), 500.0D,
                    "here is another button", comboBoxValues[1] } };

    private final ComboBox<String> comboBox;

    private boolean initializingComboBoxValue;

    private Button button;

    private Button button2;

    private Button button3;

    private Button button4;

    private Button button5;

    private boolean hidden = false;

    private NativeSelect<String> nativeSelect;

    private ComboBox<String> comboBox2;

    private ComponentsExample componentsExample;

    public TestComponentFactory() {
        testWorkbook = new XSSFWorkbook();
        final Sheet sheet = getTestWorkbook().createSheet("Custom Components");
        Row lastRow = sheet.createRow(100);
        lastRow.createCell(100, Cell.CELL_TYPE_BOOLEAN).setCellValue(true);
        sheet.setColumnWidth(0, 7800);
        sheet.setColumnWidth(1, 3800);
        sheet.setColumnWidth(2, 6100);
        sheet.setColumnWidth(3, 7200);
        sheet.setColumnWidth(4, 5900);
        sheet.setColumnWidth(5, 8200);

        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i);
            row.setHeightInPoints(28F);
            for (int j = 0; j < data[0].length; j++) {
                Cell cell = row.createCell(j);
                Object value = data[i][j];
                if (i == 0 || j == 0 || j == 4 || j == 5) {
                    // string cells
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                } else if (j == 2 || j == 3) {
                    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                }
                final DataFormat format = getTestWorkbook().createDataFormat();
                if (value != null) {
                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Double) {
                        cell.setCellValue((Double) value);
                        CellStyle style = sheet.getWorkbook().createCellStyle();
                        style.setDataFormat(format.getFormat("0000.0"));
                        cell.setCellStyle(style);
                    } else if (value instanceof Integer) {
                        cell.setCellValue(((Integer) value).intValue());
                        CellStyle style = sheet.getWorkbook().createCellStyle();
                        style.setDataFormat(format.getFormat("0.0"));
                        cell.setCellStyle(style);
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    } else if (value instanceof Date) {
                        cell.setCellValue((Date) value);
                        CellStyle dateStyle = sheet.getWorkbook()
                                .createCellStyle();
                        dateStyle
                                .setDataFormat(format.getFormat("m/d/yy h:mm"));
                        cell.setCellStyle(dateStyle);
                    } else if (value instanceof Calendar) {
                        cell.setCellValue((Calendar) value);
                        CellStyle dateStyle = sheet.getWorkbook()
                                .createCellStyle();
                        dateStyle.setDataFormat(format.getFormat("d m yyyy"));
                        cell.setCellStyle(dateStyle);
                    }
                } // null sells don't get a value
            }
        }
        Row row5 = sheet.createRow(5);
        row5.setHeightInPoints(28F);
        row5.createCell(0).setCellValue(
                "This cell has a value, and a component (label)");
        row5.createCell(1).setCellValue("This cell has a value, and a button");
        Cell cell2 = row5.createCell(2);
        cell2.setCellValue("This cell has a value and button, and is locked.");
        CellStyle lockedCellStyle = sheet.getWorkbook().createCellStyle();
        lockedCellStyle.setLocked(true);
        cell2.setCellStyle(lockedCellStyle);
        Row row6 = sheet.createRow(6);
        row6.setHeightInPoints(28F);
        comboBox = new ComboBox<>();
        comboBox.setItems(comboBoxValues);
        comboBox.addValueChangeListener(event -> {
            if (!initializingComboBoxValue) {
                String s = comboBox.getValue();
                CellReference cr = getSpreadsheet()
                        .getSelectedCellReference();
                Cell cell = getSpreadsheet().getCell(cr.getRow(),
                        cr.getCol());
                if (cell != null) {
                    cell.setCellValue(s);
                    getSpreadsheet().refreshCells(cell);
                }
            }
        });
        comboBox.setWidth("100%");
        // comboBox.setWidth("100px");

        dateField.addValueChangeListener(event-> {
                CellReference selectedCellReference = getSpreadsheet()
                        .getSelectedCellReference();
                Cell cell = getSpreadsheet().getCell(
                        selectedCellReference.getRow(),
                        selectedCellReference.getCol());
                try {
                    Date oldValue = cell.getDateCellValue();
                    Date value =Date.from(dateField.getValue().atStartOfDay(ZoneId.systemDefault()).toInstant());
                    if (oldValue != null && !oldValue.equals(value)) {
                        cell.setCellValue(value);
                        getSpreadsheet().refreshCells(cell);
                    }
                } catch (IllegalStateException ise) {
                    ise.printStackTrace();
                } catch (NullPointerException npe) {
                    npe.printStackTrace();
                }
        });
        checkBox.addValueChangeListener(event -> {

            CellReference selectedCellReference = getSpreadsheet()
                    .getSelectedCellReference();
            Cell cell = getSpreadsheet().getCell(
                    selectedCellReference.getRow(),
                    selectedCellReference.getCol());
            try {
                Boolean value = checkBox.getValue();
                Boolean oldValue = cell.getBooleanCellValue();
                if (value != oldValue) {
                    cell.setCellValue(value);
                    getSpreadsheet().refreshCells(cell);
                }
            } catch (IllegalStateException ise) {
                ise.printStackTrace();
            }
        });
    }

    public TestComponentFactory(ComponentsExample componentsExample) {
        this();
        this.componentsExample = componentsExample;
    }

    @Override
    public Component getCustomEditorForCell(Cell cell, int rowIndex,
            int columnIndex, Spreadsheet spreadsheet, Sheet sheet) {
        if (spreadsheet.getActiveSheetIndex() == 0) {
            if (rowIndex == 0 || rowIndex > 3) {
                return null;
            }
            if (1 == columnIndex) { // boolean
                return checkBox;
            } else if (2 == columnIndex) { // date
                return dateField;
            } else if (3 == columnIndex) { // numeric
                return null;
            } else if (4 == columnIndex) { // button
                return new Button("Button " + (++counter),
                        new Button.ClickListener() {

                            @Override
                            public void buttonClick(ClickEvent event) {
                                Notification
                                        .show("Clicked button inside sheet");
                            }
                        });
            } else if (5 == columnIndex) { // combobox
                return comboBox;
            }
        }
        return null;
    }

    @Override
    public void onCustomEditorDisplayed(Cell cell, int rowIndex,
            int columnIndex, Spreadsheet spreadsheet, Sheet sheet,
            Component customEditor) {
        if (customEditor instanceof Button) {
            if (rowIndex == 3) {
                customEditor.setWidth("100%");
            } else {
                customEditor.setWidth("130px");
                customEditor.setCaption("Col " + columnIndex + " Row "
                        + rowIndex);
            }
            return;
        }
        if (customEditor.equals(comboBox)) {
            initializingComboBoxValue = true;
            String stringCellValue = cell != null ? cell.getStringCellValue()
                    : null;
            comboBox.setValue(stringCellValue);
            initializingComboBoxValue = false;
        }

        if (cell != null) {
            if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                ((CheckBox) customEditor).setValue(cell.getBooleanCellValue());
            } else if (customEditor instanceof DateField) {
                final String s = cell.getCellStyle().getDataFormatString();
                if (s.contains("d")) {
                    ((DateField) customEditor)
                            .setResolution(DateResolution.DAY);
                } else if (s.contains("m") || s.contains("mmm")) {
                    ((DateField) customEditor)
                            .setResolution(DateResolution.MONTH);
                } else {
                    ((DateField) customEditor)
                            .setResolution(DateResolution.YEAR);
                }

                if(cell.getDateCellValue()!=null) {
                    LocalDate date = cell.getDateCellValue().toInstant()
                            .atZone(ZoneId.systemDefault()).toLocalDate();
                    ((DateField) customEditor).setValue(date);
                }
                Format format = spreadsheet.getDataFormatter().createFormat(
                        cell);
                String pattern = null;
                if (format instanceof ExcelStyleDateFormatter) {
                    pattern = ((ExcelStyleDateFormatter) format)
                            .toLocalizedPattern();
                }
                try {
                    ((DateField) customEditor).setDateFormat(pattern);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    @Override
    public Component getCustomComponentForCell(Cell cell, final int rowIndex,
            final int columnIndex, final Spreadsheet spreadsheet,
            final Sheet sheet) {

        if (rowIndex == 5) {
            if (!hidden) {
                if (columnIndex == 0) {
                    Label label = new Label(
                            "<div style=\"text-overflow: ellipsis; overflow: hidden; white-space: nowrap;\">Row with Custom Components</div>",
                            ContentMode.HTML);
                    return label;
                }
                if (columnIndex == 1) {
                    if (button == null) {
                        button = new Button("CLICKME",
                                new Button.ClickListener() {

                                    @Override
                                    public void buttonClick(ClickEvent event) {
                                        Notification
                                                .show("Clicked button at row index "
                                                        + rowIndex
                                                        + " column index "
                                                        + columnIndex);
                                    }
                                });
                        button.setWidth("100%");
                    }
                    return button;
                }
                if (columnIndex == 2) {
                    if (button3 == null) {
                        button3 = new Button("Hide/Show rows 1-4",
                                new Button.ClickListener() {

                                    @Override
                                    public void buttonClick(ClickEvent event) {
                                        boolean hidden = !sheet.getRow(0)
                                                .getZeroHeight();
                                        spreadsheet.setRowHidden(0, hidden);
                                        spreadsheet.setRowHidden(1, hidden);
                                        spreadsheet.setRowHidden(2, hidden);
                                        spreadsheet.setRowHidden(3, hidden);
                                    }
                                });
                    }
                    return button3;
                }
                if (columnIndex == 3) {
                    if (button2 == null) {
                        button2 = new Button("Hide/Show Columns F-I",
                                new Button.ClickListener() {

                                    @Override
                                    public void buttonClick(ClickEvent event) {
                                        boolean hidden = !sheet
                                                .isColumnHidden(5);
                                        spreadsheet.setColumnHidden(5, hidden);
                                        spreadsheet.setColumnHidden(6, hidden);
                                        spreadsheet.setColumnHidden(7, hidden);
                                        spreadsheet.setColumnHidden(8, hidden);
                                    }
                                });
                    }
                    return button2;
                }
                if (columnIndex == 4) {
                    if (button4 == null) {
                        button4 = new Button("Lock/Unlock sheet",
                                new Button.ClickListener() {

                                    @Override
                                    public void buttonClick(ClickEvent event) {
                                        if (spreadsheet.getActiveSheet()
                                                .getProtect()) {
                                            spreadsheet
                                                    .setActiveSheetProtected(null);
                                        } else {
                                            spreadsheet
                                                    .setActiveSheetProtected("");
                                        }
                                    }
                                });
                    }
                    return button4;
                }
            }
            if (columnIndex == 5) {
                if (button5 == null) {
                    button5 = new Button("Hide all custom components",
                            new Button.ClickListener() {

                                @Override
                                public void buttonClick(ClickEvent event) {
                                    hidden = !hidden;
                                    spreadsheet.reloadVisibleCellContents();
                                }
                            });
                }
                return button5;
            }
        } else if (!hidden && rowIndex == 6) {
            if (columnIndex == 1) {
                if (nativeSelect == null) {
                    nativeSelect = new NativeSelect<>();
                    List<String> items = new ArrayList<>();
                    items.add("JEE");
                    nativeSelect.setItems(items);
                    nativeSelect.setHeight("100%");
                    nativeSelect.setWidth("100%");
                }
                return nativeSelect;
            } else if (columnIndex == 2) {
                if (comboBox2 == null) {
                    comboBox2 = new ComboBox<>();
                    comboBox2.setItems(comboBoxValues);
                    comboBox2.setWidth("100%");
                }
                return comboBox2;
            }
        }
        return null;
    }

    /**
     * @return the testWorkbook
     */
    public Workbook getTestWorkbook() {
        return testWorkbook;
    }

    public Spreadsheet getSpreadsheet() {
        return componentsExample.getSpreadsheet();
    }

}