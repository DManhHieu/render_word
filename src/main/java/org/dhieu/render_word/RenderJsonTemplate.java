package org.dhieu.render_word;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.IOException;

import static org.dhieu.render_word.RenderWordUtil.renderWord;

public class RenderJsonTemplate {


    public static void main(String[] args) {
        JSONObject jsonObject = new JSONObject();
        jsonObject.put("demo", "Demo render doc");
        JSONArray jsonArray = new JSONArray();
        jsonArray.put(new TableObject("STT", "Họ và tên", "Ngày sinh", "Ghi chú", "Địa chỉ").toJson());
        jsonArray.put(new TableObject("1", "Nguyen Van A", "01/02/2000", "XYZ", "ABCC").toJson());
        jsonArray.put(new TableObject("2", "Nguyen Van B", "01/02/2000", "ABC", "XXX").toJson());
        jsonArray.put(new TableObject("3", "Nguyen Van C", "01/04/2000", "XXX", "YYY").toJson());

        jsonObject.put("table_name", jsonArray);
        jsonObject.put("name", "Đoàn Hiếu");
        try {
            RenderWordUtil.renderWord(jsonObject, "demo.docx", "simple.docx");
        } catch (IOException | OpenXML4JException e) {
            e.printStackTrace();
        }
    }


    public static class TableObject {
        private String stt;
        private String name;
        private String date;
        private String note;
        private String address;

        public TableObject(String stt, String name, String date, String note, String address) {
            this.stt = stt;
            this.name = name;
            this.date = date;
            this.note = note;
            this.address = address;
        }

        public JSONObject toJson() {
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("stt", stt);
            jsonObject.put("name", name);
            jsonObject.put("date", date);
            jsonObject.put("note", note);
            jsonObject.put("address", address);
            return jsonObject;
        }

        public String getStt() {
            return stt;
        }

        public void setStt(String stt) {
            this.stt = stt;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getDate() {
            return date;
        }

        public void setDate(String date) {
            this.date = date;
        }

        public String getNote() {
            return note;
        }

        public void setAddress(String address) {
            this.address = address;
        }

        public String getAddress() {
            return address;
        }

        public void setNote(String note) {
            this.note = note;
        }
    }
}

