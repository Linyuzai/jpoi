package com.github.linyuzai.jpoi.excel.adapter;

import java.util.Arrays;
import java.util.List;

public class MultiListWriteAdapter extends LambdaFieldDataWriteAdapter {

    public MultiListWriteAdapter() {

    }

    public MultiListWriteAdapter(List<?>... dataList) {
        this(Arrays.asList(dataList));
    }

    public MultiListWriteAdapter(List<List<?>> dataList) {
        if (dataList == null) {
            throw new RuntimeException("DataList is null");
        }
        for (List<?> data : dataList) {
            addListData(data);
        }
    }

    public void addListData(List<?> dataList) {
        addListData(null, dataList, (Class<?>) null);
    }

    public void addListData(String sheetName, List<?> dataList) {
        addListData(sheetName, dataList, (Class<?>) null);
    }

    public void addListData(List<?> dataList, Class<?> cls) {
        addListData(null, dataList, cls);
    }

    public void addListData(String sheetName, List<?> dataList, Class<?> cls) {
        if (dataList == null) {
            throw new RuntimeException("DataList is null");
        }
        ListData listData = new ListData();
        listData.setDataList(dataList);
        listData.setSheetName(sheetName);
        addListData(listData, cls);
    }

    public <T> void addListData(List<?> dataList, LambdaMethod<T, ?>... lambdaMethods) {
        addListData(null, dataList, lambdaMethods);
    }

    public <T> void addListData(String sheetName, List<?> dataList, LambdaMethod<T, ?>... lambdaMethods) {
        ListData listData = new ListData();
        listData.setDataList(dataList);
        listData.setSheetName(sheetName);
        addListData(listData, lambdaMethods);
    }
}
