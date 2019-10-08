package com.github.linyuzai.jpoi.excel.value.format;

public class NumericDataFormatValue implements SupportDataFormat {

    private double value;

    private String formatValue;

    private short dataFormat;

    public NumericDataFormatValue(double value, String formatValue, short dataFormat) {
        this.value = value;
        this.formatValue = formatValue;
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
    public String getFormatValue() {
        return formatValue;
    }

    public void setFormatValue(String formatValue) {
        this.formatValue = formatValue;
    }

    @Override
    public short getDataFormat() {
        return dataFormat;
    }

    public void setDataFormat(short dataFormat) {
        this.dataFormat = dataFormat;
    }
}
