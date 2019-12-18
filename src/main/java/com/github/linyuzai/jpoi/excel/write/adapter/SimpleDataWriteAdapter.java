package com.github.linyuzai.jpoi.excel.write.adapter;

import java.util.List;

public class SimpleDataWriteAdapter extends TitleIndexDataWriteAdapter {

    public SimpleDataWriteAdapter(List<?> dataList) {
        addListData(dataList);
    }

    public SimpleDataWriteAdapter(List<?> dataList, Class<?> cls) {
        addListData(dataList, cls);
    }

    @SafeVarargs
    public <T> SimpleDataWriteAdapter(List<?> dataList, LambdaMethod<T, ?>... lambdaMethods) throws Throwable {
        addListData(dataList, lambdaMethods);
    }

    public SimpleDataWriteAdapter(String sheetName, List<?> dataList) {
        addListData(sheetName, dataList);
    }

    public SimpleDataWriteAdapter(String sheetName, List<?> dataList, Class<?> cls) {
        addListData(sheetName, dataList, cls);
    }

    @SafeVarargs
    public <T> SimpleDataWriteAdapter(String sheetName, List<?> dataList, LambdaMethod<T, ?>... lambdaMethods) throws Throwable {
        addListData(sheetName, dataList, lambdaMethods);
    }
}
