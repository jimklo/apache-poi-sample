package com.sri.jklo;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.EnumSet;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created by jklo on 11/2/16.
 */
public class StyledDocument {

    private static final Logger logger = Logger.getLogger(StyledDocument.class.getName());

    private final String styleTemplate = "/base-template.docx";

    private XWPFDocument document;

    private XWPFNumbering numbering;


    protected Map<String, XWPFAbstractNum> numberStyles = new HashMap<String, XWPFAbstractNum>();

    protected void createDocFromTemplate() {
        try {
            document = new XWPFDocument(this.getClass().getResourceAsStream(styleTemplate));

            // strip "body" content from template
            int pos = document.getBodyElements().size()-1;
            while (pos >= 0) {
                IBodyElement element = document.getBodyElements().get(pos);
                if (!EnumSet.of(BodyType.HEADER, BodyType.FOOTER).contains(element.getPartType())) {
                    boolean success = document.removeBodyElement(pos);
                    logger.log(Level.INFO, "Removed body element "+pos+": "+success);
                }
                pos--;
            }

            initNumberingStyles();

        } catch (IOException e) {
            logger.log(Level.WARNING, "Not able to load style template", e);
            document = new XWPFDocument();
        }

    }

    /**
     * first discover all the numbering styles defined in the template.
     * a bit brute force since I can't find a way to just enumerate all the
     * abstractNum's inside the numbering.xml
     */
    protected void initNumberingStyles() {
        numbering = document.getNumbering();

        BigInteger curIdx = BigInteger.ONE;
        XWPFAbstractNum abstractNum;

        while ((abstractNum = numbering.getAbstractNum(curIdx)) != null) {
            if (abstractNum != null) {
                CTString pStyle = abstractNum.getCTAbstractNum().getLvlArray(0).getPStyle();
                if (pStyle != null) {
                    numberStyles.put(pStyle.getVal(), abstractNum);
                }
            }
            curIdx = curIdx.add(BigInteger.ONE);
        }

    }


    /**
     * This creates a new num based upon the specified numberStyle
     * @param numberStyle
     * @return
     */
    private XWPFNum restartNumbering(String numberStyle) {
        XWPFAbstractNum abstractNum = numberStyles.get(numberStyle);
        BigInteger numId = numbering.addNum(abstractNum.getAbstractNum().getAbstractNumId());
        XWPFNum num = numbering.getNum(numId);
        CTNumLvl lvlOverride = num.getCTNum().addNewLvlOverride();
        lvlOverride.setIlvl(BigInteger.ZERO);
        CTDecimalNumber number = lvlOverride.addNewStartOverride();
        number.setVal(BigInteger.ONE);
        return num;
    }


    /**
     * This creates a five item list with a simple heading, using the specified style..
     * @param index
     * @param styleName
     */
    protected void createStyledNumberList(int index, String styleName) {
        XWPFParagraph p = document.createParagraph();
        XWPFRun run = p.createRun();
        run.setText(String.format("List %d: - %s", index, styleName));

        // restart numbering
        XWPFNum num = restartNumbering(styleName);

        for (int i=1; i<=5; i++) {
            XWPFParagraph p2 = document.createParagraph();

            // set the style for this paragraph
            p2.setStyle(styleName);

            // set numbering for paragraph
            p2.setNumID(num.getCTNum().getNumId());
            CTNumPr numProp = p2.getCTP().getPPr().getNumPr();
            numProp.addNewIlvl().setVal(BigInteger.ZERO);

            // set the text
            XWPFRun run2 = p2.createRun();
            run2.setText(String.format("Item #%d using '%s' style.", i, styleName));
        }

        // some whitespace
        p = document.createParagraph();
        p.createRun();

    }

    public void createReport() {
        createDocFromTemplate();

        for (int a=0; a<5; a++) {
            int i = 0;
            for (String styleName : numberStyles.keySet()) {
                createStyledNumberList(++i, styleName);
            }
        }

    }

    public void write(OutputStream outputStream) throws IOException {
        document.write(outputStream);
    }


    public static void main(String[] args) {
        File outputFile = new File("/tmp/StyledDocument.docx");

        StyledDocument doc = new StyledDocument();
        doc.createReport();

        FileOutputStream os = null;
        try {
            os = new FileOutputStream(outputFile);
            doc.write(os);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


}
