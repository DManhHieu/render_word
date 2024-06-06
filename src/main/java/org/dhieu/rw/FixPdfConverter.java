package org.dhieu.rw;

import fr.opensagres.poi.xwpf.converter.core.IXWPFConverter;
import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Writer;

public class FixPdfConverter extends PdfConverter {
    private static final IXWPFConverter<PdfOptions> INSTANCE = new FixPdfConverter();

    public static IXWPFConverter<PdfOptions> getInstance() {
        return INSTANCE;
    }

    @Override
    protected void doConvert(XWPFDocument document, OutputStream out,
                             Writer writer, PdfOptions options) throws XWPFConverterException,
            IOException {
        try {
            // PdfMapper mapper = new PdfMapper( document, out, options );

            // process content
            ByteArrayOutputStream tempOut = new ByteArrayOutputStream();
            PdfMapper mapper = new PdfMapper(document, tempOut, options, null);
            mapper.start();

            if (mapper.useTotalPageField()) {
                // process content a second time, knowing the expected page
                // number
                Integer actualPageCount = Integer
                        .valueOf(mapper.getPageCount());
                mapper = new PdfMapper(document, out, options, actualPageCount);
                mapper.start();
            } else {
                out.write(tempOut.toByteArray());
            }

        } catch (Exception e) {
            throw new XWPFConverterException(e);
        }

    }
}
