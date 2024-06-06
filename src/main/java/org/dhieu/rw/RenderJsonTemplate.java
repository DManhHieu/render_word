package org.dhieu.rw;

import com.lowagie.text.FontFactory;
import org.json.JSONObject;
import org.json.JSONTokener;

import java.io.InputStream;
import java.util.Objects;

public class RenderJsonTemplate {
    static {
        FontFactory.registerDirectory(Objects.requireNonNull(RenderJsonTemplate.class.getClassLoader().getResource("font")).getPath(), true);
    }
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

