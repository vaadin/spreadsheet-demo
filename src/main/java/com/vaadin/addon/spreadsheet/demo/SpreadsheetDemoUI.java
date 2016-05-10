package com.vaadin.addon.spreadsheet.demo;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

import javax.servlet.annotation.WebServlet;

import org.reflections.Reflections;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.SpreadsheetFactory;
import com.vaadin.addon.spreadsheet.demo.examples.FileUploadExample;
import com.vaadin.addon.spreadsheet.demo.examples.SkipFromDemo;
import com.vaadin.addon.spreadsheet.demo.examples.SpreadsheetExample;
import com.vaadin.annotations.Theme;
import com.vaadin.annotations.Title;
import com.vaadin.annotations.VaadinServletConfiguration;
import com.vaadin.data.Container;
import com.vaadin.data.Item;
import com.vaadin.data.Property.ValueChangeEvent;
import com.vaadin.data.Property.ValueChangeListener;
import com.vaadin.data.util.FilesystemContainer;
import com.vaadin.data.util.HierarchicalContainer;
import com.vaadin.server.ExternalResource;
import com.vaadin.server.FontAwesome;
import com.vaadin.server.VaadinRequest;
import com.vaadin.server.VaadinServlet;
import com.vaadin.shared.ui.MarginInfo;
import com.vaadin.ui.CssLayout;
import com.vaadin.ui.HorizontalSplitPanel;
import com.vaadin.ui.Label;
import com.vaadin.ui.Link;
import com.vaadin.ui.Tree;
import com.vaadin.ui.UI;
import com.vaadin.ui.VerticalLayout;

/**
 * Demo class for the Spreadsheet component.
 * <p>
 * You can upload any xls or xlsx file using the upload component. You can also
 * place spreadsheet files on the classpath, under the folder /testsheets/, and
 * they will be picked up in a combobox in the menu.
 * 
 * 
 */
@SuppressWarnings({ "serial", "rawtypes", "unchecked" })
@Theme("demo-theme")
@Title("Vaadin Spreadsheet Demo")
public class SpreadsheetDemoUI extends UI implements ValueChangeListener {

    @WebServlet(value = "/*", asyncSupported = true)
    @VaadinServletConfiguration(productionMode = false, ui = SpreadsheetDemoUI.class, widgetset = "com.vaadin.addon.spreadsheet.demo.DemoWidgetSet")
    public static class Servlet extends VaadinServlet {
    }

    private Tree tree;
    private HorizontalSplitPanel horizontalSplitPanel;

    public SpreadsheetDemoUI() {
        super();
        setSizeFull();
        SpreadsheetFactory.logMemoryUsage();
    }

    @Override
    protected void init(VaadinRequest request) {
        horizontalSplitPanel = new HorizontalSplitPanel();
        horizontalSplitPanel.setSplitPosition(300, Unit.PIXELS);
        horizontalSplitPanel.addStyleName("main-layout");

        final Link github = new Link("Source code on Github",
                new ExternalResource(
                        "https://github.com/vaadin/spreadsheet-demo"));
        github.setIcon(FontAwesome.GITHUB);
        github.addStyleName("link");

        setContent(new CssLayout() {
            {
                setSizeFull();
                addComponent(horizontalSplitPanel);
                addComponent(github);
            }
        });

        VerticalLayout content = new VerticalLayout();
        content.setSpacing(true);
        content.setMargin(new MarginInfo(false, false, true, false));

        Label logo = new Label("Vaadin Spreadsheet");
        logo.addStyleName("h3");
        logo.addStyleName("logo");

        tree = new Tree();
        tree.setImmediate(true);
        tree.setContainerDataSource(getContainer());
        tree.setItemCaptionPropertyId("displayName");
        tree.setNullSelectionAllowed(false);
        tree.setWidth("100%");
        tree.addValueChangeListener(this);
        content.addComponents(logo, tree);
        horizontalSplitPanel.setFirstComponent(content);

        initSelection();
    }

    private void initSelection() {
        Iterator<?> iterator = tree.getItemIds().iterator();
        if (iterator.hasNext()) {
            tree.select(iterator.next());
        }
    }

    private Container getContainer() {
        HierarchicalContainer hierarchicalContainer = new HierarchicalContainer();
        hierarchicalContainer.addContainerProperty("order", Integer.class, 1);
        hierarchicalContainer.addContainerProperty("displayName", String.class,
                "");

        Item fileItem;
        Collection<File> files = getFiles();
        for (File file : files) {
            fileItem = hierarchicalContainer.addItem(file);
            if (file.getName().equals("Loan Calculator.xlsx")) {
                fileItem.getItemProperty("order").setValue(0);
            }
            if (file.getName().equals("Embedded Charts.xlsx")) {
                fileItem.getItemProperty("order").setValue(2);
            }
            fileItem.getItemProperty("displayName").setValue(
                    file.getName().subSequence(0, file.getName().indexOf('.')));
            hierarchicalContainer.setChildrenAllowed(file, false);
        }

        Item groupItem;
        List<Class<? extends SpreadsheetExample>> examples = getExamples();
        for (Class<? extends SpreadsheetExample> class1 : examples) {
            if (class1.getAnnotation(SkipFromDemo.class) != null) {
                continue;
            }
            groupItem = hierarchicalContainer.addItem(class1);
            if (class1.getSimpleName().equals("ChartExample")) {
                groupItem.getItemProperty("order").setValue(2);
            }
            groupItem.getItemProperty("displayName")
                    .setValue(splitCamelCase(class1.getSimpleName()));
            hierarchicalContainer.setChildrenAllowed(class1, false);
        }

        boolean[] ascending = { true, true };
        hierarchicalContainer.sort(
                hierarchicalContainer.getContainerPropertyIds().toArray(),
                ascending);

        return hierarchicalContainer;
    }

    private Collection<File> getFiles() {
        File root = null;
        try {
            ClassLoader classLoader = SpreadsheetDemoUI.class.getClassLoader();
            URL resource = classLoader
                    .getResource("testsheets" + File.separator);
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
        return testSheetContainer.getItemIds();
    }

    private List<Class<? extends SpreadsheetExample>> getExamples() {
        Reflections reflections = new Reflections(
                "com.vaadin.addon.spreadsheet.demo.examples");
        List<Class<? extends SpreadsheetExample>> examples = new ArrayList<Class<? extends SpreadsheetExample>>(
                reflections.getSubTypesOf(SpreadsheetExample.class));
        Collections.sort(examples,
                new Comparator<Class<? extends SpreadsheetExample>>() {
                    @Override
                    public int compare(Class<? extends SpreadsheetExample> o1,
                            Class<? extends SpreadsheetExample> o2) {
                        String simpleName = o1.getSimpleName();
                        String simpleName2 = o2.getSimpleName();
                        return simpleName.compareTo(simpleName2);
                    }
                });
        return examples;
    }

    static String splitCamelCase(String s) {
        String replaced = s.replaceAll(
                String.format("%s|%s|%s", "(?<=[A-Z])(?=[A-Z][a-z])",
                        "(?<=[^A-Z])(?=[A-Z])", "(?<=[A-Za-z])(?=[^A-Za-z])"),
                " ");
        replaced = replaced.replaceAll("Example", "");
        return replaced.trim();
    }

    @Override
    public void valueChange(ValueChangeEvent event) {
        Object value = event.getProperty().getValue();
        if (value instanceof File || value instanceof Class) {
            open(value);
        } else {
            tree.expandItemsRecursively(value);
            if (tree.hasChildren(value)) {
                Object firstChild = tree.getChildren(value).iterator().next();
                open(firstChild);
                tree.setValue(firstChild);
            }
        }
    }

    private void open(Object value) {
        if (value instanceof File) {
            openFile((File) value);
        } else if (value instanceof Class) {
            openExample((Class) value);
        }
    }

    private void openExample(Class value) {
        try {
            SpreadsheetExample example = (SpreadsheetExample) value
                    .newInstance();
            horizontalSplitPanel.setSecondComponent(example.getComponent());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void openFile(File file) {
        try {
            Spreadsheet spreadsheet = new Spreadsheet(file);
            horizontalSplitPanel.setSecondComponent(spreadsheet);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}
