package org.dhieu.rw;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RenderWord extends BaseRenderWord {

    private final Pattern pattern = Pattern.compile("\\$\\{(.*?)}");

//    @Override
//    protected void writeOutput(XWPFDocument doc, String output) {
//        try (FileOutputStream out = new FileOutputStream(output)) {
//            doc.write(out);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

    @Override
    protected void writeOutput(XWPFDocument doc, String output) {
        try (FileOutputStream out = new FileOutputStream(output)) {
            PdfOptions pdfOptions = PdfOptions.getDefault();
            PdfConverter.getInstance().convert(doc, out, pdfOptions);
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

    @Override
    protected String toFormatValuable(String key) {
        return "${" + key + "}";
    }

    @Override
    protected String getValuableName(String text) {
        Matcher matcher = pattern.matcher(text);
        if (matcher.find()) {
            return matcher.group().replace("${", "").replace("}", "").trim();
        }
        return null;
    }

}
