package com.vaadin.addon.spreadsheet.demo;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.servlet.annotation.WebServlet;

import org.apache.commons.io.IOUtils;
import org.reflections.Reflections;

import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.addon.spreadsheet.SpreadsheetFactory;
import com.vaadin.addon.spreadsheet.demo.examples.SkipFromDemo;
import com.vaadin.addon.spreadsheet.demo.examples.SpreadsheetExample;
import com.vaadin.addon.spreadsheet.demo.helpers.FileExampleHelper;
import com.vaadin.addon.spreadsheet.demo.helpers.NavigationBarHelper;
import com.vaadin.annotations.JavaScript;
import com.vaadin.annotations.Theme;
import com.vaadin.annotations.Title;
import com.vaadin.annotations.VaadinServletConfiguration;
import com.vaadin.server.ExternalResource;
import com.vaadin.server.FontAwesome;
import com.vaadin.server.VaadinRequest;
import com.vaadin.server.VaadinServlet;
import com.vaadin.shared.ui.label.ContentMode;
import com.vaadin.ui.Alignment;
import com.vaadin.ui.Button;
import com.vaadin.ui.HorizontalLayout;
import com.vaadin.ui.HorizontalSplitPanel;
import com.vaadin.ui.Label;
import com.vaadin.ui.Link;
import com.vaadin.ui.Panel;
import com.vaadin.ui.TabSheet;
import com.vaadin.ui.TabSheet.SelectedTabChangeEvent;
import com.vaadin.ui.TabSheet.SelectedTabChangeListener;
import com.vaadin.ui.UI;
import com.vaadin.ui.VerticalLayout;
import com.vaadin.ui.themes.ValoTheme;
import com.vaadin.v7.data.Container;
import com.vaadin.v7.data.Item;
import com.vaadin.v7.data.util.FilesystemContainer;
import com.vaadin.v7.data.util.HierarchicalContainer;
import com.vaadin.v7.ui.Tree;

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
@JavaScript("prettify.js")
@Theme("demo-theme")
@Title("Vaadin Spreadsheet Demo")
public class SpreadsheetDemoUI extends UI {

    static final Properties prop = new Properties();
    static {
        try {
            // load a properties file
            prop.load(SpreadsheetDemoUI.class
                    .getResourceAsStream("config.properties"));
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    @WebServlet(value = "/*", asyncSupported = true)
    @VaadinServletConfiguration(productionMode = false, ui = SpreadsheetDemoUI.class, widgetset = "com.vaadin.addon.spreadsheet.demo.DemoWidgetSet")
    public static class Servlet extends VaadinServlet {
    }

    private Tree tree;
    private TabSheet tabSheet;
    private NavigationBarHelper navigationBarHelper;

    public SpreadsheetDemoUI() {
        super();
        setSizeFull();
        SpreadsheetFactory.logMemoryUsage();
    }

    @Override
    protected void init(VaadinRequest request) {
        HorizontalSplitPanel horizontalSplitPanel = new HorizontalSplitPanel();
        horizontalSplitPanel.setSplitPosition(300, Unit.PIXELS);
        horizontalSplitPanel.setLocked(true);
        horizontalSplitPanel.addStyleName("main-layout");

        final Link github = new Link("Source code on Github",
                new ExternalResource(
                        "https://github.com/vaadin/spreadsheet-demo"));
        github.setIcon(FontAwesome.GITHUB);
        github.addStyleName("link");

        Label version = new Label("Version " + getVersion());
        version.addStyleName("version");

        setContent(horizontalSplitPanel);

        VerticalLayout content = new VerticalLayout();
        content.setMargin(false);
        content.setSpacing(true);

        Label logo = new Label("Vaadin Spreadsheet");
        logo.addStyleName("h3");
        logo.addStyleName("logo");
        logo.setWidth(100,Unit.PERCENTAGE);
        initNavigationBarHelper();

        Link homepage = new Link("Home page", new ExternalResource(
                "https://vaadin.com/spreadsheet"));
        Link javadoc = new Link("JavaDoc", new ExternalResource(
                "http://demo.vaadin.com/javadoc/com.vaadin.addon/vaadin-spreadsheet/"
                        + getVersion() + "/"));
        Link manual = new Link(
                "Manual",
                new ExternalResource(
                        "https://vaadin.com/docs/-/part/spreadsheet/spreadsheet-overview.html"));

        HorizontalLayout links = new HorizontalLayout(homepage, javadoc, manual);
        links.setSpacing(true);
        links.addStyleName("links");

        tree = new Tree();
        tree.setImmediate(true);
        tree.setHtmlContentAllowed(true);
        tree.setContainerDataSource(getContainer());
        tree.setItemCaptionPropertyId("displayName");
        tree.setNullSelectionAllowed(false);
        tree.setWidth("100%");
        tree.addValueChangeListener(e->{
            Object value = e.getProperty().getValue();
            open(value);
        });
        for (Object itemId : tree.rootItemIds()) {
            tree.expandItem(itemId);
        }
        Panel panel = new Panel();
        panel.setContent(tree);
        panel.setSizeFull();
        panel.setStyleName("panel");

        Button feedback = new Button("Got feedback?");
        feedback.setIcon(FontAwesome.COMMENTING_O);
        feedback.addStyleName("feedback-button");
        feedback.addStyleName(ValoTheme.BUTTON_PRIMARY);
        feedback.addStyleName(ValoTheme.BUTTON_TINY);
        feedback.addClickListener(e -> {
            getUI().addWindow(new FeedbackForm());
        });

        content.setSizeFull();
        content.addComponents(logo, links, feedback, panel, github, version);
        content.setComponentAlignment(feedback, Alignment.MIDDLE_CENTER);
        content.setExpandRatio(panel, 1);

        horizontalSplitPanel.setFirstComponent(content);

        tabSheet = new TabSheet();
        tabSheet.addSelectedTabChangeListener(new SelectedTabChangeListener() {
            @Override
            public void selectedTabChange(SelectedTabChangeEvent event) {
                com.vaadin.ui.JavaScript
                        .eval("setTimeout(function(){prettyPrint();},300);");
            }
        });
        tabSheet.setSizeFull();
        tabSheet.addStyleName(ValoTheme.TABSHEET_PADDED_TABBAR);
        horizontalSplitPanel.setSecondComponent(tabSheet);

        initSelection();
    }

    static String getVersion() {
        return (String) prop.get("spreadsheet.version");
    }

    private void initSelection() {
        Iterator<?> iterator = tree.getItemIds().iterator();
        if (iterator.hasNext()) {
            tree.select(iterator.next());
        }
    }

    private void initNavigationBarHelper() {
        // add a navigation bar item for each example to the navigation bar
        // helper
        navigationBarHelper = new NavigationBarHelper();

        // add items for file examples
        navigationBarHelper.addNavigationItem("Named Ranges Chart.xlsx",
                "Basic functionality", "Edit imported Excel file with "
                        + "<br>formatting, basic formulas, and a <br>chart. "
                                + "Updates dynamically when <br> values are edited.",
                        0);
        navigationBarHelper.addNavigationItem("Formulas.xlsx",
                "Collaborative features",
                "Freeze panes, protected cells <br> and add comments", 0);
        navigationBarHelper.addNavigationItem("Simple Invoice.xlsx",
                "Simple invoice", "Use the spreadsheet for invoices", 1);
        navigationBarHelper.addNavigationItem("Embedded Charts.xlsx",
                "Embedded charts",
                "Display charts from an Excel file <br> in the spreadsheet", 3);
        navigationBarHelper.addNavigationItem("Grouping.xlsx", "Grouping",
                "Use the Excel feature for <br> grouping rows and colums", 1);

        // add items for class examples
        navigationBarHelper.addNavigationItem("BasicStylingExample",
                "Formatting", "Style your spreadsheet", 1);
        navigationBarHelper.addNavigationItem("ChartExample", "Data binding",
                "Display spreadsheet data <br> using Vaadin charts", 2);
        navigationBarHelper.addNavigationItem("ComponentsExample",
                "Use inline components",
                "Use Vaadin components within <br> a spreadsheet", 1);
        navigationBarHelper.addNavigationItem("FileUploadExample",
                "Upload Excel files", "Upload a .xlsx or .xls file", 1);
        navigationBarHelper.addNavigationItem("ReportModeExample",
                "Report mode", "Use the read only mode <br> of spreadsheet", 1);
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
            fileItem.getItemProperty("order").setValue(
                    navigationBarHelper.getOrderNumber(file.getName()));
            fileItem.getItemProperty("displayName").setValue(
                    navigationBarHelper.getDisplayName(file.getName()));
            hierarchicalContainer.setChildrenAllowed(file, false);
        }

        Item groupItem;
        List<Class<? extends SpreadsheetExample>> examples = getExamples();
        for (Class<? extends SpreadsheetExample> class1 : examples) {
            if (class1.getAnnotation(SkipFromDemo.class) != null) {
                continue;
            }
            groupItem = hierarchicalContainer.addItem(class1);
            groupItem.getItemProperty("order").setValue(
                    navigationBarHelper.getOrderNumber(class1.getSimpleName()));
            groupItem.getItemProperty("displayName").setValue(
                    navigationBarHelper.getDisplayName(class1.getSimpleName()));
            hierarchicalContainer.setChildrenAllowed(class1, false);
        }

        boolean[] ascending = { true, true };
        hierarchicalContainer.sort(hierarchicalContainer
                .getContainerPropertyIds().toArray(), ascending);

        return hierarchicalContainer;
    }

    private Collection<File> getFiles() {
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
        return testSheetContainer.getItemIds();
    }

    private List<Class<? extends SpreadsheetExample>> getExamples() {
        Reflections reflections = new Reflections(
                "com.vaadin.addon.spreadsheet.demo.examples");
        List<Class<? extends SpreadsheetExample>> examples = new ArrayList<>(
                reflections.getSubTypesOf(SpreadsheetExample.class));
        return examples;
    }

    static String splitCamelCase(String s) {
        String replaced = s.replaceAll(String.format("%s|%s|%s",
                "(?<=[A-Z])(?=[A-Z][a-z])", "(?<=[^A-Z])(?=[A-Z])",
                "(?<=[A-Za-z])(?=[^A-Za-z])"), " ");
        replaced = replaced.replaceAll("Example", "");
        return replaced.trim();
    }

    private void open(Object value) {
        tabSheet.removeAllComponents();
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
            tabSheet.addTab(example.getComponent(), "Demo");
            addResourceTab(value, value.getSimpleName() + ".java",
                    "Java Source");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void openFile(File file) {
        Spreadsheet spreadsheet = FileExampleHelper.openFile(file);
        tabSheet.addTab(spreadsheet, "Demo");
        addResourceTab(FileExampleHelper.class,
                FileExampleHelper.class.getSimpleName() + ".java",
                "Java Source");
    }

    private void addResourceTab(Class clazz, String resourceName, String tabName) {
        try {
            InputStream resourceAsStream = clazz
                    .getResourceAsStream(resourceName);
            String code = IOUtils.toString(resourceAsStream);

            Panel p = getSourcePanel(code);

            tabSheet.addTab(p, tabName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Panel getSourcePanel(String code) {
        Panel p = new Panel();
        p.setWidth("100%");
        p.setStyleName(ValoTheme.PANEL_BORDERLESS);
        code = code.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;");
        Label c = new Label("<pre class='prettyprint'>" + code + "</pre>");
        c.setContentMode(ContentMode.HTML);
        c.setSizeUndefined();
        p.setContent(c);
        return p;
    }

}
