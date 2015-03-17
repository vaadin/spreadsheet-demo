package com.vaadin.addon.spreadsheet.demo.examples;

import static org.apache.commons.lang3.StringUtils.isEmpty;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

import org.apache.poi.ss.usermodel.Cell;

import com.vaadin.addon.charts.Chart;
import com.vaadin.addon.charts.model.ChartType;
import com.vaadin.addon.charts.model.Configuration;
import com.vaadin.addon.charts.model.DataSeries;
import com.vaadin.addon.charts.model.DataSeriesItem;
import com.vaadin.addon.charts.model.Labels;
import com.vaadin.addon.charts.model.PlotOptionsPie;
import com.vaadin.addon.charts.model.Tooltip;
import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.Spreadsheet.CellValueChangeEvent;
import com.vaadin.addon.spreadsheet.Spreadsheet.CellValueChangeListener;
import com.vaadin.ui.Component;
import com.vaadin.ui.VerticalLayout;

public class ChartExample implements SpreadsheetExample {

    private VerticalLayout layout;
    private Chart chart;
    private Spreadsheet spreadsheet;

    public ChartExample() {
        layout = new VerticalLayout();
        layout.setSizeFull();

        initSpreadsheet();
        initChart();
        layout.addComponents(chart, spreadsheet);
        layout.setExpandRatio(chart, 1.0f);
        layout.setExpandRatio(spreadsheet, 1.0f);
    }

    @Override
    public Component getComponent() {
        return layout;
    }

    private void initChart() {
        chart = new Chart(ChartType.PIE);
        chart.setSizeFull();
        Configuration conf = chart.getConfiguration();
        PlotOptionsPie plotOptions = new PlotOptionsPie();
        plotOptions.setTooltip(new Tooltip());
        plotOptions.getTooltip().setEnabled(false);
        plotOptions.setAnimation(false);
        Labels labels = new Labels();
        labels.setEnabled(true);
        labels.setFormatter("''+ this.point.name +': '+ this.percentage.toFixed(2) +' %'");
        plotOptions.setDataLabels(labels);
        conf.setPlotOptions(plotOptions);

        updateChartData();

    }

    private void initSpreadsheet() {
        spreadsheet = new Spreadsheet();
        spreadsheet.addCellValueChangeListener(new CellValueChangeListener() {
            @Override
            public void onCellValueChange(CellValueChangeEvent event) {
                updateChartData();
            }
        });
        spreadsheet.createCell(0, 0,
                "Edit this spreadsheet to alter chart title and data");
        spreadsheet.createCell(1, 0, "This is chart title");
        spreadsheet.createCell(2, 0, "Category");
        spreadsheet.createCell(2, 1, "Amount");
        spreadsheet.createCell(3, 0, "Brand 1");
        spreadsheet.createCell(3, 1, 90d);
        spreadsheet.createCell(4, 0, "Brand 2");
        spreadsheet.createCell(4, 1, 7d);
        spreadsheet.createCell(5, 0, "Brand 3");
        spreadsheet.createCell(5, 1, 3d);
        spreadsheet.setColumnWidth(0, 130);
    }

    private void updateChartData() {
        int rowIndex = 3;
        Configuration conf = chart.getConfiguration();
        String oldTitle = conf.getTitle().getText();
        String newTitle = getStringValue(1, 0);
        conf.setTitle(newTitle);
        DataSeries oldSeries = null;
        if (!conf.getSeries().isEmpty()) {
            oldSeries = (DataSeries) conf.getSeries().get(0);
        }
        DataSeries series = new DataSeries();
        while (!isEmpty(getStringValue(rowIndex, 0))) {
            series.add(new DataSeriesItem(getStringValue(rowIndex, 0),
                    getNumericValue(rowIndex, 1)));
            rowIndex++;
        }
        if (oldSeries == null
                || !series.toString().equals(oldSeries.toString())
                || !newTitle.equals(oldTitle)) {
            conf.setSeries(series);
            chart.drawChart();
        }
    }

    private String getStringValue(int rowIndex, int columnIndex) {
        Cell cell = spreadsheet.getCell(rowIndex, columnIndex);
        if (cell != null) {
            cell.setCellType(Cell.CELL_TYPE_STRING);
            return cell.getStringCellValue();
        }
        return null;
    }

    private Double getNumericValue(int rowIndex, int columnIndex) {
        Cell cell = spreadsheet.getCell(rowIndex, columnIndex);
        if (cell != null
                && (cell.getCellType() == CELL_TYPE_NUMERIC || (cell
                        .getCellType() == CELL_TYPE_FORMULA && cell
                        .getCachedFormulaResultType() == CELL_TYPE_NUMERIC))) {
            return cell.getNumericCellValue();
        }
        return 0d;
    }

}
