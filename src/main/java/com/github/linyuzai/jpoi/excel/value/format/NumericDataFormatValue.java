package com.github.linyuzai.jpoi.excel.value.format;

public class NumericDataFormatValue implements SupportDataFormat {

    private double value;

    private short dataFormat;

    public NumericDataFormatValue(double value, short dataFormat) {
        this.value = value;
        this.dataFormat = dataFormat;
    }

    @Override
    public double getValue() {
        return value;
    }

    public void setValue(double value) {
        this.value = value;
    }

    @Override
    public short getDataFormat() {
        return dataFormat;
    }

    public void setDataFormat(short dataFormat) {
        this.dataFormat = dataFormat;
    }
}
