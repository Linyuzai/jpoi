package com.github.linyuzai.jpoi.excel.write.adapter;

import com.github.linyuzai.jpoi.excel.converter.WriteCommentValueConverter;
import com.github.linyuzai.jpoi.excel.value.combination.ListCombinationValue;
import com.github.linyuzai.jpoi.excel.value.comment.StringComment;

import java.util.List;

public class TitleIndexDataWriteAdapter extends MultiListWriteAdapter {

    private boolean titleEnable = true;
    private boolean indexEnable = false;

    public TitleIndexDataWriteAdapter() {
    }

    public TitleIndexDataWriteAdapter(List<?>... dataList) {
        super(dataList);
    }

    public TitleIndexDataWriteAdapter(List<List<?>> dataList) {
        super(dataList);
    }

    public boolean isTitleEnable() {
        return titleEnable;
    }

    public void setTitleEnable(boolean titleEnable) {
        this.titleEnable = titleEnable;
    }

    public boolean isIndexEnable() {
        return indexEnable;
    }

    public void setIndexEnable(boolean indexEnable) {
        this.indexEnable = indexEnable;
    }

    @Override
    public int getHeaderRowCount(int sheet) {
        return titleEnable ? 1 : 0;
    }

    @Override
    public Object getHeaderRowContent(int sheet, int row, int cell, int realRow, int realCell) {
        if (!titleEnable || cell < 0) {
            return null;
        }
        WriteField writeField = getWriteFieldList().get(sheet).get(cell);
        String fieldDescription = writeField.getFieldDescription();
        if (writeField instanceof AnnotationWriteField) {
            String tc = ((AnnotationWriteField) writeField).getTitleComment();
            if (tc != null) {
                ListCombinationValue combinationValue = new ListCombinationValue();
                combinationValue.addValue(fieldDescription);
                combinationValue.addValue(WriteCommentValueConverter.getInstance().convertValue(sheet, realRow, realCell, tc));
                return combinationValue;
            }
        }
        return fieldDescription;
    }

    @Override
    public int getHeaderCellCount(int sheet, int row) {
        return indexEnable ? 1 : 0;
    }

    @Override
    public Object getHeaderCellContent(int sheet, int row, int cell, int realRow, int realCell) {
        if (!indexEnable || row < 0) {
            return null;
        }
        return String.valueOf(row);
    }
}
