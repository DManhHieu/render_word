package org.dhieu.rw;

import com.lowagie.text.Font;
import fr.opensagres.poi.xwpf.converter.pdf.internal.elements.StylableParagraph;

public class FixStylableParagraph
        extends StylableParagraph {

    private boolean defaultLeading = true;
    private final StylableParagraph stylableParagraph;

    public FixStylableParagraph(StylableParagraph paragraph) {
        super(paragraph.getOwnerDocument(), paragraph.getParent());
        this.defaultLeading = false;
        this.stylableParagraph = paragraph;
    }

    public void adjustLeading(Font font) {
        if (font != null) {
            super.setLeading(font.getSize());
            this.stylableParagraph.setLeading(font.getSize());
            this.defaultLeading = false;
        }
    }

}
