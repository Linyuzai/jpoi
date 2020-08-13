package com.github.linyuzai.jpoi.excel.write.adapter;

import com.github.linyuzai.jpoi.common.SerializedLambda;
import com.github.linyuzai.jpoi.exception.JPoiException;

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
            throw new JPoiException("DataList is null");
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
            throw new JPoiException("DataList is null");
        }
        ListData listData = new ListData();
        listData.setDataList(dataList);
        listData.setSheetName(sheetName);
        addListData(listData, cls);
    }

    @SafeVarargs
    public final <T> void addListData(List<?> dataList, SerializedLambda.Func<T, ?>... lambdaMethods) throws Throwable {
        addListData(null, dataList, lambdaMethods);
    }

    @SafeVarargs
    public final <T> void addListData(String sheetName, List<?> dataList, SerializedLambda.Func<T, ?>... lambdaMethods) throws Throwable {
        ListData listData = new ListData();
        listData.setDataList(dataList);
        listData.setSheetName(sheetName);
        addListData(listData, lambdaMethods);
    }
}
