package org.dhieu.rw;

import org.json.JSONObject;
import org.json.JSONTokener;

import java.io.InputStream;

public class RenderJsonTemplate {

    public static void main(String[] args) {
        try {
            IRenderWord renderWord = new RenderWord();
            ClassLoader classloader = Thread.currentThread().getContextClassLoader();
            InputStream inputStream = classloader.getResourceAsStream("demo.json");
            assert inputStream != null;
            JSONObject jsonObject = new JSONObject(new JSONTokener((inputStream)));
            renderWord.render(jsonObject, "demo.docx", "demo.pdf");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

