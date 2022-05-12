package org.dhieu.render_word;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public final class RenderWordUtil {
    public static void renderWord(JSONObject jsonObject, String urlTemplate, String output) throws IOException, OpenXML4JException {
        InputStream input = readTemplate(urlTemplate);
        renderWord(jsonObject, input, output);
    }

    public static void renderWord(JSONObject jsonObject, InputStream input, String output) throws IOException {

        try (XWPFDocument doc = new XWPFDocument(input)) {
            for (XWPFTable table : doc.getTables()) {
                renderTable(table, jsonObject);
            }

            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                renderParagraph(paragraph, jsonObject);
            }

            for (XWPFHeader xwpfHeader : doc.getHeaderList()) {
                renderHeader(xwpfHeader, jsonObject);
            }
            for (XWPFFooter xwpfFooter : doc.getFooterList()) {
                renderFooter(xwpfFooter, jsonObject);
            }

            for (XWPFPictureData pictureData : doc.getAllPictures()) {
                renderPicture(pictureData, jsonObject);
            }
            writeOutput(doc, output);
        }
    }


    private static void renderFooter(XWPFFooter xwpfFooter, JSONObject jsonObject) {
    }

    private static void renderHeader(XWPFHeader xwpfHeader, JSONObject jsonObject) {
    }

    private static void writeOutput(XWPFDocument doc, String output) {
        try (FileOutputStream out = new FileOutputStream(output)) {
            doc.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static InputStream readTemplate(String urlTemplate) {
        ClassLoader classloader = Thread.currentThread().getContextClassLoader();
        return classloader.getResourceAsStream(urlTemplate);
    }

    private static void renderPicture(XWPFPictureData pictureData, JSONObject jsonObject) {
        System.out.println(pictureData);
    }

    private static void renderParagraphType(XWPFParagraph paragraph, JSONObject jsonObject) {
        if (paragraph.getRuns() != null) {
            for (XWPFRun run : paragraph.getRuns()) {
                fillObject(run, jsonObject);
            }
        }
    }

    private static void renderParagraph(XWPFParagraph paragraph, JSONObject jsonObject) {
        renderParagraphType(paragraph, jsonObject);
    }

    private static void fillObject(XWPFRun run, JSONObject jsonObject) {
        String text = run.getText(0);
        if (text != null && !text.isBlank()) {
            for (Map.Entry<String, Object> entry : jsonObject.toMap().entrySet()) {
                text = text.replace("{" + entry.getKey() + "}", entry.getValue().toString());
            }
            run.setText(text, 0);
        }
    }

    private static String getValuableName(String text) {
        return text.substring(text.indexOf("{") + 1, text.indexOf("}"));
    }

    private static TableFormat getFormatTable(XWPFTable table) {
        TableFormat tableFormat = new TableFormat();
        List<XWPFTableCell> formatTableCell = table.getRows().get(0).getTableCells();
        String format = formatTableCell.get(0).getText();
        tableFormat.setTableName(getValuableName(format));
        tableFormat.setFields(new ArrayList<>());
        String field0String = format.substring(format.indexOf("}") + 1);
        String field = getValuableName(field0String);
        tableFormat.getFields().add(field);
        for (int i = 1; i < formatTableCell.size(); i++) {
            String cellText = table.getRows().get(0).getTableCells().get(i).getText();
            if (cellText != null && !cellText.isBlank()) {
                tableFormat.getFields().add(getValuableName(cellText));
            }
        }
        return tableFormat;
    }

    private static void renderTable(XWPFTable table, JSONObject jsonObject) {
        TableFormat tableFormat = getFormatTable(table);
        int tableSize = table.getRows().size();
        JSONArray jsonArray = jsonObject.getJSONArray(tableFormat.getTableName());
        for (int i = 0; i < Math.max(jsonArray.length(), tableSize); i++) {
            if (i < jsonArray.length() && tableSize > i && (i == 0 || checkEmptyCell(table.getRows().get(i)))) {
                fillTableCell(table.getRows().get(i), jsonArray.getJSONObject(i), tableFormat);
            } else if (i < jsonArray.length() && (tableSize <= i || !checkEmptyCell(table.getRows().get(i)))) {
                XWPFTableRow newRow = new XWPFTableRow((CTRow) table.getRows().get(i - 1).getCtRow().copy(), table);
                fillTableCell(newRow, jsonArray.getJSONObject(i), tableFormat);
                table.addRow(newRow, i);
                tableSize++;
            }
            fillObjectTableCell(table.getRows().get(i), jsonObject);
        }

    }

    private static void fillObjectTableCell(XWPFTableRow row, JSONObject jsonObject) {
        for (XWPFTableCell tableCell : row.getTableCells()) {
            for (XWPFRun run : tableCell.getParagraphs().get(0).getRuns()) {
                fillObject(run, jsonObject);
            }
        }
    }

    private static boolean checkEmptyCell(XWPFTableRow row) {
        for (XWPFTableCell tableCell : row.getTableCells()) {
            if (!tableCell.getText().isBlank()) {
                return false;
            }
        }
        return true;
    }

    private static void fillTableCell(XWPFTableRow row, JSONObject object, TableFormat tableFormat) {
        for (int i = 0; i < tableFormat.getFields().size(); i++) {
            row.getTableCells().get(i).getParagraphs().get(0).getRuns().get(0).setText((String) object.get(tableFormat.getFields().get(i)), 0);
        }
    }

    private static class TableFormat {
        private String tableName;
        private List<String> fields;

        public TableFormat() {
        }

        public TableFormat(String tableName, List<String> fields) {
            this.fields = fields;
            this.tableName = tableName;
        }

        public String getTableName() {
            return tableName;
        }

        public void setFields(List<String> fields) {
            this.fields = fields;
        }

        public void setTableName(String tableName) {
            this.tableName = tableName;
        }

        public List<String> getFields() {
            return fields;
        }
    }
}
