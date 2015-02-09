package com.vaadin.addon.spreadsheet.demo;

import static com.vaadin.ui.themes.ValoTheme.THEME_NAME;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import javax.servlet.annotation.WebServlet;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.Spreadsheet.ProtectedEditEvent;
import com.vaadin.addon.spreadsheet.Spreadsheet.ProtectedEditListener;
import com.vaadin.addon.spreadsheet.Spreadsheet.SelectedSheetChangeEvent;
import com.vaadin.addon.spreadsheet.Spreadsheet.SelectedSheetChangeListener;
import com.vaadin.addon.spreadsheet.Spreadsheet.SelectionChangeEvent;
import com.vaadin.addon.spreadsheet.Spreadsheet.SelectionChangeListener;
import com.vaadin.addon.spreadsheet.SpreadsheetComponentFactory;
import com.vaadin.addon.spreadsheet.SpreadsheetFactory;
import com.vaadin.addon.spreadsheet.action.SpreadsheetDefaultActionHandler;
import com.vaadin.annotations.Theme;
import com.vaadin.annotations.VaadinServletConfiguration;
import com.vaadin.data.Property.ValueChangeEvent;
import com.vaadin.data.Property.ValueChangeListener;
import com.vaadin.data.util.FilesystemContainer;
import com.vaadin.event.Action.Handler;
import com.vaadin.server.FileDownloader;
import com.vaadin.server.FileResource;
import com.vaadin.server.FontAwesome;
import com.vaadin.server.VaadinRequest;
import com.vaadin.server.VaadinServlet;
import com.vaadin.shared.ui.colorpicker.Color;
import com.vaadin.ui.Alignment;
import com.vaadin.ui.Button;
import com.vaadin.ui.Button.ClickEvent;
import com.vaadin.ui.CheckBox;
import com.vaadin.ui.ColorPicker;
import com.vaadin.ui.ComboBox;
import com.vaadin.ui.HorizontalLayout;
import com.vaadin.ui.Notification;
import com.vaadin.ui.UI;
import com.vaadin.ui.Upload;
import com.vaadin.ui.Upload.Receiver;
import com.vaadin.ui.Upload.SucceededEvent;
import com.vaadin.ui.Upload.SucceededListener;
import com.vaadin.ui.VerticalLayout;
import com.vaadin.ui.Window;
import com.vaadin.ui.components.colorpicker.ColorChangeEvent;
import com.vaadin.ui.components.colorpicker.ColorChangeListener;

/**
 * Demo class for the Spreadsheet component.
 * <p>
 * You can upload any xls or xlsx file using the upload component. You can also
 * place spreadsheet files on the classpath, under the folder /testsheets/, and
 * they will be picked up in a combobox in the menu.
 * 
 * 
 */
@SuppressWarnings("serial")
@Theme(THEME_NAME)
public class SpreadsheetDemoUI extends UI implements Receiver {

    @WebServlet(value = "/*", asyncSupported = true)
    @VaadinServletConfiguration(productionMode = false, ui = SpreadsheetDemoUI.class, widgetset = "com.vaadin.addon.spreadsheet.demo.DemoWidgetSet")
    public static class Servlet extends VaadinServlet {
    }

    VerticalLayout layout = new VerticalLayout();

    Upload upload = new Upload();
    private File previousFile = null;

    private Button save;
    private Button download;
    private File uploadedFile;
    private ComboBox openTestSheetSelect;
    private CheckBox gridlines;
    private CheckBox rowColHeadings;

    Spreadsheet spreadsheet;
    private SelectionChangeListener selectionChangeListener;
    private SpreadsheetComponentFactory spreadsheetFieldFactory;
    private SelectedSheetChangeListener selectedSheetChangeListener;
    private final Handler spreadsheetActionHandler = new SpreadsheetDefaultActionHandler();

    public SpreadsheetDemoUI() {
        super();
        SpreadsheetFactory.logMemoryUsage();
    }

    @Override
    protected void init(VaadinRequest request) {
        SpreadsheetFactory.logMemoryUsage();
        setContent(layout);

        buildOptions();

        selectionChangeListener = new SelectionChangeListener() {

            @Override
            public void onSelectionChange(SelectionChangeEvent event) {
                printSelectionChangeEventContents(event);
            }
        };

        selectedSheetChangeListener = new SelectedSheetChangeListener() {

            @Override
            public void onSelectedSheetChange(SelectedSheetChangeEvent event) {
                gridlines.setValue(event.getNewSheet().isDisplayGridlines());
                rowColHeadings.setValue(event.getNewSheet()
                        .isDisplayRowColHeadings());
            }
        };

        spreadsheetFieldFactory = new TestComponentFactory();

    }

    private void buildOptions() {
        VerticalLayout menuBar = new VerticalLayout();
        menuBar.setSpacing(true);

        HorizontalLayout options = new HorizontalLayout();
        options.setSpacing(true);

        layout.setMargin(true);
        layout.setSpacing(true);

        layout.setSizeFull();

        gridlines = new CheckBox("grid lines");
        gridlines.setImmediate(true);

        rowColHeadings = new CheckBox("headers");
        rowColHeadings.setImmediate(true);

        gridlines.addValueChangeListener(new ValueChangeListener() {

            @Override
            public void valueChange(ValueChangeEvent event) {
                Boolean display = (Boolean) event.getProperty().getValue();

                if (spreadsheet != null) {
                    spreadsheet.setGridlinesVisible(display);
                }
            }
        });

        rowColHeadings.addValueChangeListener(new ValueChangeListener() {

            @Override
            public void valueChange(ValueChangeEvent event) {
                Boolean display = (Boolean) event.getProperty().getValue();

                if (spreadsheet != null) {
                    spreadsheet.setRowColHeadingsVisible(display);
                }
            }
        });

        Button newSpreadsheetButton = new Button("New Spreadsheet",
                new Button.ClickListener() {

                    @Override
                    public void buttonClick(ClickEvent event) {
                        createNewSheet();
                    }
                });

        Button newSpreadsheetInWindowButton = new Button("New in Window",
                new Button.ClickListener() {

                    @Override
                    public void buttonClick(ClickEvent event) {
                        createNewSheetInWindow();
                    }
                });

        File root = null;
        try {
            ClassLoader classLoader = SpreadsheetDemoUI.class.getClassLoader();
            URL resource = classLoader.getResource("testsheets"
                    + File.separator);
            if (resource != null) {
                root = new File(resource.toURI());
            }
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
        FilesystemContainer testSheetContainer = new FilesystemContainer(root);
        testSheetContainer.setRecursive(false);
        testSheetContainer.setFilter(new FilenameFilter() {

            @Override
            public boolean accept(File dir, String name) {
                if (name != null
                        && (name.endsWith(".xls") || name.endsWith(".xlsx"))) {
                    return true;
                } else {
                    return false;
                }
            }
        });
        openTestSheetSelect = new ComboBox(null, testSheetContainer);
        openTestSheetSelect.setPageLength(0);
        openTestSheetSelect.setImmediate(true);
        openTestSheetSelect.setItemCaptionPropertyId("Name");
        openTestSheetSelect.setItemIconPropertyId("Icon");
        openTestSheetSelect.addValueChangeListener(new ValueChangeListener() {

            @Override
            public void valueChange(ValueChangeEvent event) {
                Object value = openTestSheetSelect.getValue();
                if (value instanceof File) {
                    loadFile((File) value);
                }
            }
        });
        save = new Button("Save", new Button.ClickListener() {

            @Override
            public void buttonClick(ClickEvent event) {
                if (spreadsheet != null) {
                    saveFile();
                }
            }
        });
        save.setEnabled(false);
        download = new Button("Download");
        download.setEnabled(false);

        Button customComponentTest = new Button("Custom Cell Editors",
                new Button.ClickListener() {

                    @Override
                    public void buttonClick(ClickEvent event) {
                        createEditorTestSheet();
                    }

                });
        upload.setReceiver(this);

        options.addComponent(newSpreadsheetButton);
        options.addComponent(newSpreadsheetInWindowButton);
        options.addComponent(customComponentTest);
        options.addComponent(openTestSheetSelect);
        options.addComponent(upload);
        options.addComponent(new Button("Close", new Button.ClickListener() {

            @Override
            public void buttonClick(ClickEvent event) {
                if (spreadsheet != null) {
                    SpreadsheetFactory.logMemoryUsage();
                    layout.removeComponent(spreadsheet);
                    spreadsheet = null;
                    SpreadsheetFactory.logMemoryUsage();
                }
            }
        }));

        HorizontalLayout sheetOptions = new HorizontalLayout();
        sheetOptions.setSpacing(true);
        sheetOptions.addComponent(save);
        sheetOptions.addComponent(download);

        upload.setImmediate(true);
        upload.addSucceededListener(new SucceededListener() {

            @Override
            public void uploadSucceeded(SucceededEvent event) {
                loadFile(uploadedFile);
            }
        });

        Button boldButton = new Button(FontAwesome.BOLD);
        Button italicButton = new Button(FontAwesome.ITALIC);
        ColorPicker backgroundColor = new ColorPicker("Background Color");
        backgroundColor.setIcon(FontAwesome.SQUARE);
        backgroundColor.setCaption("Background Color");
        ColorPicker fontColor = new ColorPicker("Font Color");
        fontColor.setIcon(FontAwesome.FONT);
        fontColor.setCaption("Font Color");
        boldButton.addClickListener(new Button.ClickListener() {

            @Override
            public void buttonClick(ClickEvent event) {
                if (spreadsheet != null) {
                    List<Cell> cellsToRefresh = new ArrayList<Cell>();
                    for (CellReference cellRef : spreadsheet
                            .getSelectedCellReferences()) {
                        Cell cell = getOrCreateCell(cellRef);
                        CellStyle style = cloneStyle(cell);
                        Font font = cloneFont(style);
                        font.setBold(!font.getBold());
                        style.setFont(font);
                        cell.setCellStyle(style);

                        cellsToRefresh.add(cell);
                    }
                    spreadsheet.refreshCells(cellsToRefresh);
                }
            }
        });

        italicButton.addClickListener(new Button.ClickListener() {

            @Override
            public void buttonClick(ClickEvent event) {
                if (spreadsheet != null) {
                    List<Cell> cellsToRefresh = new ArrayList<Cell>();
                    for (CellReference cellRef : spreadsheet
                            .getSelectedCellReferences()) {
                        Cell cell = getOrCreateCell(cellRef);
                        CellStyle style = cloneStyle(cell);
                        Font font = cloneFont(style);
                        font.setItalic(!font.getItalic());
                        style.setFont(font);
                        cell.setCellStyle(style);
                        cellsToRefresh.add(cell);
                    }
                    spreadsheet.refreshCells(cellsToRefresh);
                }
            }
        });

        backgroundColor.addColorChangeListener(new ColorChangeListener() {

            @Override
            public void colorChanged(ColorChangeEvent event) {
                Color newColor = event.getColor();
                if (spreadsheet != null && newColor != null) {
                    List<Cell> cellsToRefresh = new ArrayList<Cell>();
                    for (CellReference cellRef : spreadsheet
                            .getSelectedCellReferences()) {
                        Cell cell = getOrCreateCell(cellRef);
                        Workbook workbook = spreadsheet.getWorkbook();
                        if (workbook instanceof XSSFWorkbook) {
                            XSSFCellStyle style = (XSSFCellStyle) cloneStyle(cell);
                            XSSFColor color = new XSSFColor(java.awt.Color
                                    .decode(newColor.getCSS()));
                            style.setFillForegroundColor(color);
                            cell.setCellStyle(style);
                        } else {
                            CellStyle style = cloneStyle(cell);
                            style.setFillForegroundColor(getSimilarColorIndex(
                                    newColor));
                            cell.setCellStyle(style);
                        }
                        cellsToRefresh.add(cell);
                    }
                    spreadsheet.refreshCells(cellsToRefresh);
                }
            }
        });

        fontColor.addColorChangeListener(new ColorChangeListener() {

            @Override
            public void colorChanged(ColorChangeEvent event) {
                Color newColor = event.getColor();
                if (spreadsheet != null && newColor != null) {
                    List<Cell> cellsToRefresh = new ArrayList<Cell>();
                    for (CellReference cellRef : spreadsheet
                            .getSelectedCellReferences()) {
                        Cell cell = getOrCreateCell(cellRef);
                        Workbook workbook = spreadsheet.getWorkbook();
                        if (workbook instanceof XSSFWorkbook) {
                            XSSFCellStyle style = (XSSFCellStyle) cloneStyle(cell);
                            XSSFColor color = new XSSFColor(java.awt.Color
                                    .decode(newColor.getCSS()));
                            XSSFFont font = (XSSFFont) cloneFont(style);
                            font.setColor(color);
                            style.setFont(font);
                            cell.setCellStyle(style);
                        } else {
                            CellStyle style = cloneStyle(cell);
                            Font font = cloneFont(style);
                            font.setColor(getSimilarColorIndex(newColor));
                            style.setFont(font);
                            cell.setCellStyle(style);
                        }
                        cellsToRefresh.add(cell);
                    }
                    spreadsheet.refreshCells(cellsToRefresh);
                }
            }
        });

        HorizontalLayout styleOptions = new HorizontalLayout();
        styleOptions.setSpacing(true);
        styleOptions.addComponents(gridlines, rowColHeadings);
        styleOptions.addComponent(boldButton);
        styleOptions.addComponent(italicButton);
        styleOptions.addComponent(fontColor);
        styleOptions.addComponent(backgroundColor);
        styleOptions.setComponentAlignment(gridlines, Alignment.MIDDLE_CENTER);
        styleOptions.setComponentAlignment(rowColHeadings,
                Alignment.MIDDLE_CENTER);
        menuBar.addComponent(options);
        menuBar.addComponent(styleOptions);
        layout.addComponent(menuBar);

    }

    private Cell getOrCreateCell(CellReference cellRef) {
        Cell cell = spreadsheet.getCell(cellRef.getRow(), cellRef.getCol());
        if (cell == null) {
            cell = spreadsheet.createCell(cellRef.getRow(), cellRef.getCol(),
                    "");
        }
        return cell;
    }

    private CellStyle cloneStyle(Cell cell) {
        CellStyle style = spreadsheet.getWorkbook().createCellStyle();
        style.cloneStyleFrom(cell.getCellStyle());
        return style;
    }

    private Font cloneFont(CellStyle cellstyle) {
        Font newFont = spreadsheet.getWorkbook().createFont();
        Font originalFont = spreadsheet.getWorkbook().getFontAt(
                cellstyle.getFontIndex());
        if (originalFont != null) {
            newFont.setBold(originalFont.getBold());
            newFont.setItalic(originalFont.getItalic());
            newFont.setFontHeight(originalFont.getFontHeight());
            newFont.setUnderline(originalFont.getUnderline());
            newFont.setStrikeout(originalFont.getStrikeout());
            if (originalFont instanceof XSSFFont) {
                XSSFFont originalXFont = (XSSFFont) originalFont;
                XSSFFont newXFont = (XSSFFont) newFont;
                newXFont.setColor(originalXFont.getXSSFColor());
            } else {
                newFont.setColor(originalFont.getColor());
            }
        }
        return newFont;
    }

    protected void createNewSheet() {
        if (spreadsheet == null) {
            spreadsheet = new Spreadsheet();
            spreadsheet.addSelectionChangeListener(selectionChangeListener);
            spreadsheet
                    .addSelectedSheetChangeListener(selectedSheetChangeListener);
            spreadsheet.addActionHandler(spreadsheetActionHandler);

            layout.addComponent(spreadsheet);
            layout.setExpandRatio(spreadsheet, 1.0f);
        } else {
            spreadsheet.reset();
        }
        spreadsheet.setSpreadsheetComponentFactory(null);
        save.setEnabled(true);
        previousFile = null;
        openTestSheetSelect.setValue(null);

        gridlines.setValue(spreadsheet.isGridlinesVisible());
        rowColHeadings.setValue(spreadsheet.isRowColHeadingsVisible());
    }

    protected void createNewSheetInWindow() {
        Spreadsheet spreadsheet = new Spreadsheet();
        Window w = new Window("new Spreadsheet", spreadsheet);
        w.setWidth("50%");
        w.setHeight("50%");
        w.center();
        w.setModal(true);

        addWindow(w);
    }

    protected void saveFile() {
        try {
            if (previousFile != null) {
                int i = previousFile.getName().lastIndexOf(".xls");
                String fileName = previousFile.getName().substring(0, i)
                        + ("(1)") + previousFile.getName().substring(i);
                previousFile = spreadsheet.write(fileName);
            } else {
                previousFile = spreadsheet.write("workbook1");
            }
            download.setEnabled(true);
            FileResource resource = new FileResource(previousFile);
            FileDownloader fileDownloader = new FileDownloader(resource);
            fileDownloader.extend(download);
            previousFile.deleteOnExit();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    protected void createEditorTestSheet() {
        if (spreadsheet == null) {
            spreadsheet = new Spreadsheet(
                    ((TestComponentFactory) spreadsheetFieldFactory)
                            .getTestWorkbook());
            spreadsheet.setSpreadsheetComponentFactory(spreadsheetFieldFactory);
            spreadsheet.addActionHandler(spreadsheetActionHandler);
            layout.addComponent(spreadsheet);

            layout.setExpandRatio(spreadsheet, 1.0f);
        } else {
            spreadsheet
                    .setWorkbook(((TestComponentFactory) spreadsheetFieldFactory)
                            .getTestWorkbook());
            spreadsheet.setSpreadsheetComponentFactory(spreadsheetFieldFactory);
        }

        gridlines.setValue(spreadsheet.isGridlinesVisible());
        rowColHeadings.setValue(spreadsheet.isRowColHeadingsVisible());
    }

    private void printSelectionChangeEventContents(SelectionChangeEvent event) {

        Set<CellReference> allSelectedCells = event.getAllSelectedCells();
        spreadsheet.setInfoLabelValue(allSelectedCells.size()
                + " selected cells");

    }

    private void loadFile(File file) {
        try {
            if (spreadsheet == null) {
                spreadsheet = new Spreadsheet(file);
                spreadsheet.addSelectionChangeListener(selectionChangeListener);
                spreadsheet
                        .addSelectedSheetChangeListener(selectedSheetChangeListener);
                spreadsheet.addActionHandler(spreadsheetActionHandler);
                spreadsheet
                        .addProtectedEditListener(new ProtectedEditListener() {

                            @Override
                            public void writeAttempted(ProtectedEditEvent event) {
                                Notification
                                        .show("This cell is protected and cannot be changed");
                            }
                        });
                layout.addComponent(spreadsheet);
                layout.setExpandRatio(spreadsheet, 1.0f);
            } else {
                if (previousFile == null
                        || !previousFile.getAbsolutePath().equals(
                                file.getAbsolutePath())) {
                    spreadsheet.read(file);
                }
            }
            spreadsheet.setSpreadsheetComponentFactory(null);
            previousFile = file;
            save.setEnabled(true);
            download.setEnabled(false);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    @Override
    public OutputStream receiveUpload(final String filename, String mimeType) {

        try {
            File file = new File(filename);
            file.deleteOnExit();
            uploadedFile = file;
            FileOutputStream fos = new FileOutputStream(uploadedFile);
            return fos;
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return null;
    }

    private short getSimilarColorIndex(Color newColor) {
        HSSFPalette palette = ((HSSFWorkbook) spreadsheet.getWorkbook())
                .getCustomPalette();
        HSSFColor color = palette.findSimilarColor(newColor.getRed(),
                newColor.getGreen(), newColor.getBlue());
        return color.getIndex();
    }

}
