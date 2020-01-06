package com.github.linyuzai.jpoi.excel.post;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.concurrent.ExecutorService;

public class HttpURLConnectionPostProcessor extends HttpPostProcessor {

    public HttpURLConnectionPostProcessor() {
        super();
    }

    public HttpURLConnectionPostProcessor(ExecutorService executorService) {
        super(executorService);
    }

    @Override
    public void processPostValue(PostValue postValue) throws Throwable {
        if (postValue instanceof UrlPostValueHolder) {
            UrlPostValueHolder holder = (UrlPostValueHolder) postValue;
            URL url = new URL(holder.getSource());
            HttpURLConnection connection = getHttpURLConnection(url);
            connection.connect();
            InputStream is = connection.getInputStream();
            byte[] bytes = getContent(is);
            connection.disconnect();
            ValueConverter vc = holder.getValueConverter();
            holder.setValue(vc == null ? bytes : vc.convertValue(holder.getSheetIndex(), holder.getRowIndex(), holder.getCellIndex(), bytes));
        }
    }

    public HttpURLConnection getHttpURLConnection(URL url) throws IOException {
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setConnectTimeout(6 * 1000);
        connection.setReadTimeout(6 * 1000);
        connection.setDoOutput(true);
        connection.setDoInput(true);
        connection.setUseCaches(false);
        connection.setRequestMethod("GET");
        //可设置请求头
        connection.setRequestProperty("Content-Type", "application/octet-stream");
        connection.setRequestProperty("Connection", "Keep-Alive");
        connection.setRequestProperty("Charset", "UTF-8");
        return connection;
    }
}
