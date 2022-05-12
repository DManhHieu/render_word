package org.dhieu.rw;

import org.apache.commons.io.IOUtils;
import org.json.JSONObject;

import java.io.InputStream;

public class RenderJsonTemplate {


    public static void main(String[] args) {

        try {
            IRenderWord renderWord = new RenderWord();
            ClassLoader classloader = Thread.currentThread().getContextClassLoader();
            InputStream inputStream = classloader.getResourceAsStream("demo.json");
            JSONObject jsonObject = null;
            assert inputStream != null;
            jsonObject = new JSONObject(IOUtils.toString(inputStream));
            renderWord.render(jsonObject, "demo.docx", "simple.docx");
        } catch (Exception e) {
            System.out.println("ERROR");
            System.out.println(e.getMessage());
        }
    }
}

