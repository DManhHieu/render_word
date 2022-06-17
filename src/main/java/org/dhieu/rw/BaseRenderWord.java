package org.dhieu.rw;

import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public abstract class BaseRenderWord implements IRenderWord {
    public void render(JSONObject jsonObject, String urlTemplate, String output) throws IOException {
        InputStream input = readTemplate(urlTemplate);
        render(jsonObject, input, output);
    }

    public void render(JSONObject jsonObject, InputStream input, String output) throws IOException {

        XWPFDocument doc = new XWPFDocument(input);
        for (XWPFTable table : doc.getTables()) {
            renderTable(table, jsonObject);
        }

        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            renderParagraph(paragraph, jsonObject);
        }

        for (XWPFHeader xwpfHeader : doc.getHeaderList()) {
            renderHeaderFooter(xwpfHeader, jsonObject);
        }
        for (XWPFFooter xwpfFooter : doc.getFooterList()) {
            renderHeaderFooter(xwpfFooter, jsonObject);
        }

//        for (XWPFPictureData pictureData : doc.getAllPictures()) {
//            renderPicture(pictureData, jsonObject);
//        }
        writeOutput(doc, output);

    }

    protected abstract void writeOutput(XWPFDocument doc, String output);

    protected abstract InputStream readTemplate(String urlTemplate);

    protected abstract String getValuableName(String text);

    protected abstract InputStream getImage(String url) throws IOException;

    protected abstract String toFormatValuable(String key);


    private <T extends XWPFHeaderFooter> void renderHeaderFooter(T xwpfFooter, JSONObject jsonObject) throws IOException {
        for (XWPFParagraph paragraph : xwpfFooter.getParagraphs()) {
            renderParagraph(paragraph, jsonObject);
        }
        for (XWPFTable table : xwpfFooter.getTables()) {
            renderTable(table, jsonObject);
        }
//        for (XWPFPictureData pictureData : xwpfFooter.getAllPictures()) {
//            renderPicture(pictureData, jsonObject);
//        }
    }

    private void renderPicture(XWPFPictureData pictureData, JSONObject jsonObject) throws IOException {
        if (jsonObject.get(pictureData.getFileName()) != null && !jsonObject.get(pictureData.getFileName()).toString().isBlank()) {
            try (
                    InputStream inputStream = getImage(jsonObject.get(pictureData.getFileName()).toString());
                    OutputStream outputStream = pictureData.getPackagePart().getOutputStream();
            ) {
                byte[] buffer = new byte[2048];
                int lenth;
                while ((lenth = inputStream.read(buffer)) > 0) {
                    outputStream.write(buffer, 0, lenth);
                }
            }
        }
    }

    private void renderParagraph(XWPFParagraph paragraph, JSONObject jsonObject) {
        if (paragraph.getRuns() != null) {
            for (XWPFRun run : paragraph.getRuns()) {
                fillObject(run, jsonObject);
            }
        }
    }

    private void fillObject(XWPFRun run, JSONObject jsonObject) {
        String text = run.getText(0);
        if (text != null && !text.isBlank()) {
            for (Map.Entry<String, Object> entry : jsonObject.toMap().entrySet()) {
                text = text.replace(toFormatValuable(entry.getKey()), entry.getValue().toString());
            }
            run.setText(text, 0);
        }
    }

    private TableFormat getFormatTable(XWPFTable table) {
        TableFormat tableFormat = new TableFormat();
        List<XWPFTableCell> formatTableCell = table.getRows().get(0).getTableCells();
        String format = formatTableCell.get(0).getText();
        String tableName = getValuableName(format);
        if (tableName == null) {
            return null;
        }
        tableFormat.setTableName(tableName);
        tableFormat.setFields(new ArrayList<>());
        String field0String = format.substring(format.indexOf(tableFormat.getTableName()) + tableFormat.getTableName().length());
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

    private void renderTable(XWPFTable table, JSONObject jsonObject) {
        TableFormat tableFormat = getFormatTable(table);
        if (tableFormat == null) {
            return;
        }
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

    private void fillObjectTableCell(XWPFTableRow row, JSONObject jsonObject) {
        for (XWPFTableCell tableCell : row.getTableCells()) {
            for (XWPFRun run : tableCell.getParagraphs().get(0).getRuns()) {
                fillObject(run, jsonObject);
            }
        }
    }

    private boolean checkEmptyCell(XWPFTableRow row) {
        for (XWPFTableCell tableCell : row.getTableCells()) {
            if (!tableCell.getText().isBlank()) {
                return false;
            }
        }
        return true;
    }

    private void fillTableCell(XWPFTableRow row, JSONObject object, TableFormat tableFormat) {
        for (int i = 0; i < tableFormat.getFields().size(); i++) {
            row.getTableCells().get(i).getParagraphs().get(0).getRuns().get(0).setText((String) object.get(tableFormat.getFields().get(i)), 0);
            for (int j = 1; j < row.getTableCells().get(i).getParagraphs().get(0).getRuns().size(); j++) {
                row.getTableCells().get(i).getParagraphs().get(0).getRuns().get(j).setText("", 0);
            }
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
