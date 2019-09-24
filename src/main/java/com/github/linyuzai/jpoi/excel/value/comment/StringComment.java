package com.github.linyuzai.jpoi.excel.value.comment;

import com.github.linyuzai.jpoi.excel.value.ClientAnchorValue;

public class StringComment extends ClientAnchorValue implements SupportComment {

    private String comment;

    public StringComment(String comment) {
        this.comment = comment;
    }

    public String getComment() {
        return comment;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }
}
