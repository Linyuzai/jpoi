package com.github.linyuzai.jpoi.excel.post;

import com.github.linyuzai.jpoi.excel.converter.ValueConverter;
import com.github.linyuzai.jpoi.excel.holder.UrlPostValueHolder;
import com.github.linyuzai.jpoi.excel.value.post.PostValue;
import org.apache.http.HttpEntity;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpUriRequest;
import org.apache.http.impl.client.HttpClients;

import java.util.concurrent.ExecutorService;

public class HttpClientPostProcessor extends HttpPostProcessor {

    private HttpClient httpClient;

    public HttpClientPostProcessor() {
        super();
    }

    public HttpClientPostProcessor(HttpClient httpClient) {
        this(httpClient, null);
        this.httpClient = httpClient;
    }

    public HttpClientPostProcessor(ExecutorService executorService) {
        this(null, executorService);
    }

    public HttpClientPostProcessor(HttpClient httpClient, ExecutorService executorService) {
        super(executorService);
        this.httpClient = httpClient;
    }

    @Override
    public void processPostValue(PostValue postValue) throws Throwable {
        if (postValue instanceof UrlPostValueHolder) {
            UrlPostValueHolder holder = (UrlPostValueHolder) postValue;
            HttpClient client = getExecuteHttpClient(holder);
            HttpUriRequest request = getHttpUriRequest(holder);
            CloseableHttpResponse response = (CloseableHttpResponse) client.execute(request);
            HttpEntity entity = response.getEntity();
            if (entity == null) {
                return;
            }
            byte[] bytes = getContent(entity.getContent());
            ValueConverter vc = holder.getValueConverter();
            holder.setValue(vc == null ? bytes : vc.convertValue(holder.getSheetIndex(), holder.getRowIndex(), holder.getCellIndex(), bytes));
        }
    }

    public HttpUriRequest getHttpUriRequest(UrlPostValueHolder holder) {
        return new HttpGet(holder.getSource());
    }

    public HttpClient getExecuteHttpClient(UrlPostValueHolder holder) {
        return httpClient == null ? HttpClients.createDefault() : httpClient;
    }

    public HttpClient getHttpClient() {
        return httpClient;
    }

    public void setHttpClient(HttpClient httpClient) {
        this.httpClient = httpClient;
    }
}
