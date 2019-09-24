package com.github.linyuzai.jpoi.excel.value.comment;

import com.github.linyuzai.jpoi.excel.value.ClientAnchorValue;
import org.apache.poi.ss.usermodel.RichTextString;

public class RichTextStringComment extends ClientAnchorValue implements SupportComment {

    private RichTextString richTextString;

    public RichTextStringComment(RichTextString richTextString) {
        this.richTextString = richTextString;
    }

    public RichTextString getRichTextString() {
        return richTextString;
    }

    public void setRichTextString(RichTextString richTextString) {
        this.richTextString = richTextString;
    }
}
