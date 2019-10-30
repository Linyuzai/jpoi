package com.github.linyuzai.jpoi.excel.value.error;

public class ByteError implements SupportErrorValue {

    private byte errorValue;

    public ByteError(byte errorValue) {
        this.errorValue = errorValue;
    }

    @Override
    public byte getErrorValue() {
        return errorValue;
    }

    //@Override
    public void setErrorValue(byte value) {
        this.errorValue = value;
    }
}
