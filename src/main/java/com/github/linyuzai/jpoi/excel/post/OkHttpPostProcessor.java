package com.github.linyuzai.jpoi.excel.post;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;
import okhttp3.*;

import java.util.concurrent.ExecutorService;

public class OkHttpPostProcessor extends HttpPostProcessor {

    private OkHttpClient okHttpClient;

    public OkHttpPostProcessor() {
        super();
    }

    public OkHttpPostProcessor(OkHttpClient okHttpClient) {
        this(okHttpClient, null);
    }

    public OkHttpPostProcessor(ExecutorService executorService) {
        this(null, executorService);
    }

    public OkHttpPostProcessor(OkHttpClient okHttpClient, ExecutorService executorService) {
        super(executorService);
        this.okHttpClient = okHttpClient;
    }

    @Override
    public void processPostValue(PostValue postValue) throws Throwable {
        if (postValue instanceof UrlPostValueHolder) {
            UrlPostValueHolder holder = (UrlPostValueHolder) postValue;
            OkHttpClient client = getExecuteOkHttpClient(holder);
            Request request = getRequest(holder);
            Response response = client.newCall(request).execute();
            ResponseBody body = response.body();
            if (body == null) {
                return;
            }
            byte[] bytes = body.bytes();
            ValueConverter vc = holder.getValueConverter();
            holder.setValue(vc == null ? bytes : vc.convertValue(holder.getSheetIndex(), holder.getRowIndex(), holder.getCellIndex(), bytes));
        }
    }

    public Request getRequest(UrlPostValueHolder holder) {
        return new Request.Builder()
                .url(holder.getSource())
                .get()
                .build();
    }

    public OkHttpClient getExecuteOkHttpClient(UrlPostValueHolder holder) {
        return okHttpClient == null ? new OkHttpClient() : okHttpClient;
    }

    public OkHttpClient getOkHttpClient() {
        return okHttpClient;
    }

    public void setOkHttpClient(OkHttpClient okHttpClient) {
        this.okHttpClient = okHttpClient;
    }
}
