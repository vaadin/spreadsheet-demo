package com.vaadin.addon.spreadsheet.demo.examples;

import static com.vaadin.shared.ui.colorpicker.Color.BLACK;
import static com.vaadin.shared.ui.colorpicker.Color.WHITE;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.Spreadsheet.SelectionChangeEvent;
import com.vaadin.addon.spreadsheet.Spreadsheet.SelectionChangeListener;
import com.vaadin.server.FontAwesome;
import com.vaadin.shared.ui.colorpicker.Color;
import com.vaadin.ui.Button;
import com.vaadin.ui.Button.ClickEvent;
import com.vaadin.ui.Button.ClickListener;
import com.vaadin.ui.ColorPicker;
import com.vaadin.ui.Component;
import com.vaadin.ui.HorizontalLayout;
import com.vaadin.ui.VerticalLayout;
import com.vaadin.ui.components.colorpicker.ColorChangeEvent;
import com.vaadin.ui.components.colorpicker.ColorChangeListener;
import com.vaadin.ui.themes.ValoTheme;

public class BasicStylingExample
        implements SpreadsheetExample, SelectionChangeListener {

    private static final long serialVersionUID = 6330513476932056681L;

    private VerticalLayout layout;
    private HorizontalLayout stylingToolbar;
    private Spreadsheet spreadsheet;
    private ColorPicker backgroundColor;
    private ColorPicker fontColor;

    public BasicStylingExample() {
        layout = new VerticalLayout();
        layout.setSizeFull();

        initStyleToolbar();
        initSpreadsheet();
        layout.addComponents(stylingToolbar, spreadsheet);
        layout.setExpandRatio(stylingToolbar, 0f);
        layout.setExpandRatio(spreadsheet, 1.0f);
    }

    @Override
    public Component getComponent() {
        return layout;
    }

    private void initSpreadsheet() {
        spreadsheet = new Spreadsheet();
        spreadsheet.addSelectionChangeListener(this);

        Font fontBoldExample = spreadsheet.getWorkbook().createFont();
        fontBoldExample.setBold(true);
        CellStyle fontBoldExampleStyle = spreadsheet.getWorkbook()
                .createCellStyle();
        fontBoldExampleStyle.setFillBackgroundColor(HSSFColor.YELLOW.index);
        fontBoldExampleStyle.setFont(fontBoldExample);
        Cell fontExampleCell = spreadsheet.createCell(0, 0,
                "Click the 'B' button in the top left corner to toggle bold font on and off.");
        fontExampleCell.setCellStyle(fontBoldExampleStyle);

        CellStyle backgroundColorStyle = spreadsheet.getWorkbook()
                .createCellStyle();
        backgroundColorStyle.setFillBackgroundColor(HSSFColor.YELLOW.index);
        Cell backgroundExampleCell = spreadsheet.createCell(2, 0,
                "Click the 'Background Color' button to select and change the background color of a cell.");
        backgroundExampleCell.setCellStyle(backgroundColorStyle);

        Font fontColorExample = spreadsheet.getWorkbook().createFont();
        fontColorExample.setColor(HSSFColor.LIGHT_BLUE.index);
        CellStyle fontColorExampleStyle = spreadsheet.getWorkbook()
                .createCellStyle();
        fontColorExampleStyle.setFillBackgroundColor(HSSFColor.YELLOW.index);
        fontColorExampleStyle.setFont(fontColorExample);
        Cell fontColorExampleCell = spreadsheet.createCell(4, 0,
                "Click the 'Font Color' button to select and change the font color of a cell.");
        fontColorExampleCell.setCellStyle(fontColorExampleStyle);

        Cell cell;
        for (int i = 0; i <= 4; i = i + 2) {
            for (int j = 1; j <= 9; j++) {
                cell = spreadsheet.createCell(i, j, "");
                cell.setCellStyle(backgroundColorStyle);
            }
        }

        spreadsheet.refreshCells(fontExampleCell, backgroundExampleCell,
                fontColorExampleCell);

    }

    private void initStyleToolbar() {
        stylingToolbar = new HorizontalLayout();
        Button boldButton = new Button(FontAwesome.BOLD);
        boldButton.addClickListener(new ClickListener() {
            @Override
            public void buttonClick(ClickEvent event) {
                updateSelectedCellsBold();
            }
        });
        backgroundColor = new ColorPicker();
        backgroundColor.setCaption("Background Color");
        backgroundColor.addColorChangeListener(new ColorChangeListener() {
            @Override
            public void colorChanged(ColorChangeEvent event) {
                updateSelectedCellsBackgroundColor(event.getColor());
            }
        });
        fontColor = new ColorPicker();
        fontColor.setCaption("Font Color");
        fontColor.addColorChangeListener(new ColorChangeListener() {
            @Override
            public void colorChanged(ColorChangeEvent event) {
                updateSelectedCellsFontColor(event.getColor());
            }
        });
        stylingToolbar.addComponents(boldButton, backgroundColor, fontColor);
        boldButton.addStyleName(ValoTheme.BUTTON_BORDERLESS);
        backgroundColor.addStyleName(ValoTheme.BUTTON_BORDERLESS);
        fontColor.addStyleName(ValoTheme.BUTTON_BORDERLESS);
    }

    private void updateSelectedCellsBold() {
        if (spreadsheet != null) {
            List<Cell> cellsToRefresh = new ArrayList<Cell>();
            for (CellReference cellRef : spreadsheet
                    .getSelectedCellReferences()) {
                // Obtain Cell using CellReference
                Cell cell = getOrCreateCell(cellRef);
                // Clone Cell CellStyle
                CellStyle style = cloneStyle(cell);
                // Clone CellStyle Font
                Font font = cloneFont(style);
                // Toggle current bold state
                font.setBold(!font.getBold());
                style.setFont(font);
                cell.setCellStyle(style);

                cellsToRefresh.add(cell);
            }
            // Update all edited cells
            spreadsheet.refreshCells(cellsToRefresh);
        }
    }

    private void updateSelectedCellsBackgroundColor(Color newColor) {
        if (spreadsheet != null && newColor != null) {
            List<Cell> cellsToRefresh = new ArrayList<Cell>();
            for (CellReference cellRef : spreadsheet
                    .getSelectedCellReferences()) {
                // Obtain Cell using CellReference
                Cell cell = getOrCreateCell(cellRef);
                // Clone Cell CellStyle
                // This cast an only be done when using .xlsx files
                XSSFCellStyle style = (XSSFCellStyle) cloneStyle(cell);
                XSSFColor color = new XSSFColor(
                        java.awt.Color.decode(newColor.getCSS()));
                // Set new color value
                style.setFillForegroundColor(color);
                cell.setCellStyle(style);

                cellsToRefresh.add(cell);
            }
            // Update all edited cells
            spreadsheet.refreshCells(cellsToRefresh);
        }
    }

    private void updateSelectedCellsFontColor(Color newColor) {
        if (spreadsheet != null && newColor != null) {
            List<Cell> cellsToRefresh = new ArrayList<Cell>();
            for (CellReference cellRef : spreadsheet
                    .getSelectedCellReferences()) {
                Cell cell = getOrCreateCell(cellRef);
                // Workbook workbook = spreadsheet.getWorkbook();
                XSSFCellStyle style = (XSSFCellStyle) cloneStyle(cell);
                XSSFColor color = new XSSFColor(
                        java.awt.Color.decode(newColor.getCSS()));
                XSSFFont font = (XSSFFont) cloneFont(style);
                font.setColor(color);
                style.setFont(font);
                cell.setCellStyle(style);
                cellsToRefresh.add(cell);
            }
            // Update all edited cells
            spreadsheet.refreshCells(cellsToRefresh);
        }
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
        CellStyle newStyle = spreadsheet.getWorkbook().createCellStyle();
        newStyle.cloneStyleFrom(cell.getCellStyle());
        return newStyle;
    }

    private Font cloneFont(CellStyle cellstyle) {
        Font newFont = spreadsheet.getWorkbook().createFont();
        Font originalFont = spreadsheet.getWorkbook()
                .getFontAt(cellstyle.getFontIndex());
        if (originalFont != null) {
            newFont.setBold(originalFont.getBold());
            newFont.setItalic(originalFont.getItalic());
            newFont.setFontHeight(originalFont.getFontHeight());
            newFont.setUnderline(originalFont.getUnderline());
            newFont.setStrikeout(originalFont.getStrikeout());
            // This cast an only be done when using .xlsx files
            XSSFFont originalXFont = (XSSFFont) originalFont;
            XSSFFont newXFont = (XSSFFont) newFont;
            newXFont.setColor(originalXFont.getXSSFColor());
        }
        return newFont;
    }

    @Override
    public void onSelectionChange(SelectionChangeEvent event) {
        CellReference selectedCell = event.getSelectedCellReference();
        Cell cell = spreadsheet.getCell(selectedCell.getRow(),
                selectedCell.getCol());
        backgroundColor.setColor(WHITE);
        fontColor.setColor(BLACK);
        if (cell != null) {
            // This cast an only be done when using .xlsx files
            XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
            if (style != null) {
                XSSFFont font = style.getFont();
                if (font != null) {
                    XSSFColor xssfFontColor = font.getXSSFColor();
                    if (xssfFontColor != null) {
                        fontColor.setColor(convertColor(xssfFontColor));
                    }
                }
                XSSFColor foregroundColor = style.getFillForegroundColorColor();
                if (foregroundColor != null) {
                    backgroundColor.setColor(convertColor(foregroundColor));
                }
            }
        }
    }

    private Color convertColor(XSSFColor foregroundColor) {
        byte[] argb = foregroundColor.getARgb();
        return new Color(byteToInt(argb[1]), byteToInt(argb[2]),
                byteToInt(argb[3]), byteToInt(argb[0]));
    }

    private int byteToInt(byte byteValue) {
        return byteValue & 0xFF;
    }

}
