package org.apache.poi.xwpf.usermodel;

import java.util.List;

/**
 * This is a utility class so that I can get access to the protected fields within XWPFNumbering.
 * Created by jklo on 11/3/16.
 */
public class NumberingUtil {

    private final XWPFNumbering numbering;

    public NumberingUtil(XWPFNumbering numbering) {
        this.numbering = numbering;
    }

    public List<XWPFAbstractNum> getAbstractNums() {
        return numbering.abstractNums;
    }

    public List<XWPFNum> getNums() {
        return numbering.nums;
    }

    public XWPFNumbering getNumbering() {
        return numbering;
    }

}
