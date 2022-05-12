package org.dhieu.rw;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

public final class RenderWord extends BaseRenderWord {
    @Override
    protected void writeOutput(XWPFDocument doc, String output) {
        try (FileOutputStream out = new FileOutputStream(output)) {
            doc.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    protected InputStream readTemplate(String urlTemplate) {
        ClassLoader classloader = Thread.currentThread().getContextClassLoader();
        return classloader.getResourceAsStream(urlTemplate);
    }

    @Override
    protected InputStream getImage(String url) throws IOException {
        return new URL(url).openStream();
    }

}
