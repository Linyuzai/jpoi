package com.github.linyuzai.jpoi.excel.value.error;

public class IntError implements SupportErrorValue {

    private int intErrorValue;

    public IntError(int intErrorValue) {
        this.intErrorValue = intErrorValue;
    }

    @Override
    public byte getErrorValue() {
        return (byte) intErrorValue;
    }

    //@Override
    public void setErrorValue(byte value) {
        this.intErrorValue = value;
    }
}
