package com.kinnarastudio.kecakplugins.datalist.export;

import org.apache.commons.lang.StringEscapeUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.*;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.*;
import org.joget.commons.util.LogUtil;
import org.joget.commons.util.SecurityUtil;
import org.joget.commons.util.StringUtil;
import org.joget.commons.util.TimeZoneUtil;
import org.joget.plugin.base.PluginManager;
import org.joget.workflow.util.WorkflowUtil;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.util.*;

public class SelectionDatalistExcelExport extends DataListActionDefault {
    private static final String LABEL = "Selection Datalist Excel Export";

    @Override
    public DataListActionResult executeAction(DataList dataList, String[] rowKeys) {
        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }

        boolean includeAllWhenNone = "true".equalsIgnoreCase(getPropertyString("includeAllWhenNoneSelected"));
        if ((rowKeys == null || rowKeys.length == 0) && !includeAllWhenNone) {
            DataListActionResult result = new DataListActionResult();
            result.setType(DataListActionResult.TYPE_REDIRECT);
            result.setUrl("REFERER");
            result.setMessage("Please select at least one row to export.");
            return result;
        }

        try {
            exportToExcel(dataList, rowKeys);
        } catch (Throwable t) {
            LogUtil.error(getClassName(), t, "Error exporting selected rows to Excel: " + t.getMessage());
            DataListActionResult result = new DataListActionResult();
            result.setType(DataListActionResult.TYPE_ERROR);
            result.setMessage("Export error: " + t.getMessage());
            return result;
        }

        return null;
    }

    protected void exportToExcel(DataList dataList, String[] rowKeys) throws Exception {
        // Fallback for StAX in OSGi environments
        if (System.getProperty("javax.xml.stream.XMLEventFactory") == null) {
            try {
                Class.forName("com.sun.xml.internal.stream.events.XMLEventFactoryImpl");
                System.setProperty("javax.xml.stream.XMLEventFactory", "com.sun.xml.internal.stream.events.XMLEventFactoryImpl");
            } catch (Throwable ignored) {}
        }

        ClassLoader originalClassLoader = Thread.currentThread().getContextClassLoader();
        try {
            Thread.currentThread().setContextClassLoader(SelectionDatalistExcelExport.class.getClassLoader());

            List<DataListColumn> columns = getExportColumns(dataList);
            List<Object> rows = getSelectedRows(dataList, rowKeys);

            XSSFWorkbook wb = new XSSFWorkbook();
            try {
                String sheetName = getPropertyString("sheetName");
                if (sheetName == null || sheetName.trim().isEmpty()) {
                    sheetName = dataList.getName();
                    if (sheetName == null || sheetName.trim().isEmpty()) {
                        sheetName = "Sheet1";
                    }
                }
                sheetName = WorkbookUtil.createSafeSheetName(sheetName);
                XSSFSheet sheet = wb.createSheet(sheetName);

                int rowNum = 0;
                int colNum = 0;

                // Header style from CustomExcelHssfView
                XSSFCellStyle headerStyle = wb.createCellStyle();
                headerStyle.setFillPattern(XSSFCellStyle.FINE_DOTS);
                headerStyle.setFillBackgroundColor(HSSFColor.BLUE_GREY.index);
                XSSFFont bold = wb.createFont();
                bold.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                bold.setColor(HSSFColor.WHITE.index);
                headerStyle.setFont(bold);

                // Create header row
                XSSFRow headerRow = sheet.createRow(rowNum++);
                for (DataListColumn col : columns) {
                    String label = col.getLabel();
                    if (label == null || label.trim().isEmpty()) {
                        label = capitalize(col.getName());
                    }
                    XSSFCell cell = headerRow.createCell(colNum++);
                    cell.setCellValue(new XSSFRichTextString(label));
                    cell.setCellStyle(headerStyle);
                }

                // Create data rows
                for (Object row : rows) {
                    XSSFRow xlsRow = sheet.createRow(rowNum++);
                    colNum = 0;
                    for (DataListColumn col : columns) {
                        Object val = evaluateRowValue(row, col.getName());
                        Object formatted = formatColumnValue(dataList, col, row, val);
                        XSSFCell cell = xlsRow.createCell(colNum++);
                        writeCell(formatted, cell);
                    }
                }

                // Auto size columns
                for (int i = 0; i < columns.size(); i++) {
                    sheet.autoSizeColumn((short) i);
                }

                // Stream Excel directly to response
                HttpServletResponse response = WorkflowUtil.getHttpServletResponse();
                if (response != null) {
                    String filename = getPropertyString("filename");
                    if (filename == null || filename.trim().isEmpty()) {
                        filename = dataList.getId();
                        if (filename == null || filename.trim().isEmpty()) {
                            filename = "export";
                        }
                    }
                    if (!filename.toLowerCase().endsWith(".xlsx")) {
                        filename += ".xlsx";
                    }

                    String encodedFilename = URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
                    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    response.setHeader("Content-Disposition", "attachment; filename=\"" + filename + "\"; filename*=UTF-8''" + encodedFilename);

                    wb.write(response.getOutputStream());
                    response.getOutputStream().flush();
                }
            } finally {
                try {
                    wb.close();
                } catch (Exception ignored) {}
            }
        } finally {
            Thread.currentThread().setContextClassLoader(originalClassLoader);
        }
    }

    protected List<DataListColumn> getExportColumns(DataList dataList) {
        List<DataListColumn> columns = new ArrayList<DataListColumn>();
        DataListColumn[] allColumns = dataList.getColumns();
        if (allColumns != null) {
            for (DataListColumn col : allColumns) {
                if (!col.isPermitted()) {
                    continue;
                }
                if (col.isHidden()) {
                    if (!"true".equalsIgnoreCase(col.getPropertyString("include_export"))) {
                        continue;
                    }
                } else {
                    if ("true".equalsIgnoreCase(col.getPropertyString("exclude_export"))) {
                        continue;
                    }
                }
                columns.add(col);
            }
        }
        return columns;
    }

    protected List<Object> getSelectedRows(DataList dataList, String[] rowKeys) {
        List<Object> resultRows = new ArrayList<Object>();
        String primaryKey = (dataList.getBinder() != null && dataList.getBinder().getPrimaryKeyColumnName() != null)
                ? dataList.getBinder().getPrimaryKeyColumnName()
                : "id";

        if (rowKeys == null || rowKeys.length == 0) {
            DataListCollection all = dataList.getRows(DataList.MAXIMUM_PAGE_SIZE, 0);
            if (all != null) {
                resultRows.addAll(all);
            }
            return resultRows;
        }

        Set<String> selectedKeySet = new LinkedHashSet<String>(Arrays.asList(rowKeys));
        Map<String, Object> foundRows = new LinkedHashMap<String, Object>();

        // 1. Check current page rows first
        DataListCollection currentRows = dataList.getRows();
        if (currentRows != null) {
            for (Object row : currentRows) {
                Object val = evaluateRowValue(row, primaryKey);
                if (val != null && selectedKeySet.contains(val.toString())) {
                    foundRows.put(val.toString(), row);
                }
            }
        }

        // 2. If some keys are not on current page (e.g. multi-page selection), query binder
        if (foundRows.size() < selectedKeySet.size() && dataList.getBinder() != null) {
            Set<String> missingKeys = new LinkedHashSet<String>(selectedKeySet);
            missingKeys.removeAll(foundRows.keySet());

            DataListFilterQueryObject filter = new DataListFilterQueryObject();
            filter.setOperator("AND");
            String colName = dataList.getBinder().getColumnName(primaryKey);
            if (colName == null || colName.isEmpty()) {
                colName = primaryKey;
            }
            StringBuilder placeholders = new StringBuilder();
            for (int i = 0; i < missingKeys.size(); i++) {
                if (i > 0) placeholders.append(",");
                placeholders.append("?");
            }
            filter.setQuery(colName + " IN (" + placeholders + ")");
            filter.setValues(missingKeys.toArray(new String[0]));
            dataList.addFilterQueryObject(filter);

            DataListCollection extraRows = dataList.getRows(missingKeys.size(), 0);
            if (extraRows != null) {
                for (Object row : extraRows) {
                    Object val = evaluateRowValue(row, primaryKey);
                    if (val != null && missingKeys.contains(val.toString())) {
                        foundRows.put(val.toString(), row);
                    }
                }
            }
        }

        // Preserve the selection order
        for (String key : rowKeys) {
            if (foundRows.containsKey(key)) {
                resultRows.add(foundRows.get(key));
            }
        }

        return resultRows;
    }

    protected Object evaluateRowValue(Object row, String propertyName) {
        if (row == null || propertyName == null) {
            return null;
        }
        Object value = null;
        if (row instanceof Map) {
            Map<?, ?> map = (Map<?, ?>) row;
            value = map.get(propertyName);
            if (value == null) {
                value = map.get(propertyName.toLowerCase());
            }
        }
        if (value == null) {
            try {
                String capitalized = capitalize(propertyName);
                try {
                    Method m = row.getClass().getMethod("get" + capitalized);
                    value = m.invoke(row);
                } catch (NoSuchMethodException e) {
                    Method m = row.getClass().getMethod("is" + capitalized);
                    value = m.invoke(row);
                }
            } catch (Exception ignored) {}
        }
        if (value instanceof Date) {
            value = TimeZoneUtil.convertToTimeZone((Date) value, null, AppUtil.getAppDateFormat());
        }
        return value;
    }

    protected Object formatColumnValue(DataList dataList, DataListColumn column, Object row, Object value) {
        Object result = value;
        if (result instanceof String) {
            result = SecurityUtil.decrypt(result.toString());
        }

        Collection<DataListColumnFormat> formats = column.getFormats();
        if (formats != null && !formats.isEmpty()) {
            for (DataListColumnFormat format : formats) {
                if (format != null) {
                    result = format.format(dataList, column, row, result);
                }
            }
            if (result != null) {
                String str = result.toString();
                str = StringUtil.stripAllHtmlTag(str);
                str = unescapeHtml(str);
                return str;
            }
        }
        return result;
    }

    protected String unescapeHtml(String text) {
        if (text == null || text.isEmpty()) {
            return text;
        }
        try {
            return StringEscapeUtils.unescapeHtml(text);
        } catch (Throwable t) {
            return text.replace("&amp;", "&")
                    .replace("&lt;", "<")
                    .replace("&gt;", ">")
                    .replace("&quot;", "\"")
                    .replace("&#39;", "'")
                    .replace("&nbsp;", " ");
        }
    }

    protected void writeCell(Object value, XSSFCell cell) {
        if (value instanceof Number) {
            Number num = (Number) value;
            cell.setCellValue(num.doubleValue());
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else {
            cell.setCellValue(new XSSFRichTextString(escapeColumnValue(value)));
        }
    }

    protected String escapeColumnValue(Object rawValue) {
        if (rawValue == null) {
            return "";
        }
        String returnString = rawValue.toString();
        try {
            returnString = StringEscapeUtils.escapeJava(returnString.trim());
            returnString = returnString.replace("\\t", "    ");
            returnString = returnString.replace("\\r", " ");
            returnString = StringEscapeUtils.unescapeJava(returnString);
            return returnString;
        } catch (Throwable t) {
            return returnString.replace("\t", "    ").replace("\r\n", "\n").replace("\r", " ");
        }
    }

    private String capitalize(String str) {
        if (str == null || str.isEmpty()) {
            return str;
        }
        return Character.toUpperCase(str.charAt(0)) + str.substring(1);
    }

    @Override
    public String getConfirmation() {
        return getPropertyString("confirmation");
    }

    @Override
    public String getHref() {
        return getPropertyString("href");
    }

    @Override
    public String getHrefColumn() {
        return getPropertyString("hrefColumn");
    }

    @Override
    public String getHrefParam() {
        return getPropertyString("hrefParam");
    }

    @Override
    public String getLinkLabel() {
        String label = getPropertyString("label");
        if (label == null || label.isEmpty()) {
            label = "Export Selected to Excel";
        }
        return label;
    }

    @Override
    public String getTarget() {
        return "post";
    }

    @Override
    public String getClassName() {
        return getClass().getName();
    }

    @Override
    public String getLabel() {
        return LABEL;
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/SelectionDatalistExcelExport.json", null, false, null);
    }

    @Override
    public String getDescription() {
        return "";
    }

    @Override
    public String getName() {
        return getLabel();
    }

    @Override
    public String getVersion() {
        PluginManager pluginManager = (PluginManager) AppUtil.getApplicationContext().getBean("pluginManager");
        ResourceBundle resourceBundle = pluginManager.getPluginMessageBundle(getClassName(), "/messages/BuildNumber");
        return resourceBundle.getString("buildNumber");
    }
}
