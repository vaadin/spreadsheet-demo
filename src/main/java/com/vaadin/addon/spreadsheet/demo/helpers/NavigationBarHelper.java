package com.vaadin.addon.spreadsheet.demo.helpers;

import java.util.HashMap;
import java.util.Map;

public class NavigationBarHelper {

    private Map<String, String> displayNames;
    private Map<String, String> discriptions;
    private Map<String, Integer> orderNumbers;

    public NavigationBarHelper() {
        displayNames = new HashMap<String, String>();
        discriptions = new HashMap<String, String>();
        orderNumbers = new HashMap<String, Integer>();
    }

    public void addNavigationItem(String exampleName, String displayName,
            String discription, int orderNumber) {
        displayNames.put(exampleName, displayName);
        discriptions.put(exampleName, discription);
        orderNumbers.put(exampleName, orderNumber);
    }

    public String getDisplayName(String exampleName) {
        return displayNames.get(exampleName) + "<div class=\"description\">"
                + discriptions.get(exampleName) + "</div>";
    }

    public String getDiscription(String exampleName) {
        return discriptions.get(exampleName);
    }

    public int getOrderNumber(String exampleName) {
        return orderNumbers.get(exampleName);
    }

}
