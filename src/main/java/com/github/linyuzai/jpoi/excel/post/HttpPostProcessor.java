package com.github.linyuzai.jpoi.excel.post;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.ExecutorService;

public abstract class HttpPostProcessor extends ThreadPoolPostProcessor {

    public HttpPostProcessor() {
        super();
    }

    public HttpPostProcessor(ExecutorService executorService) {
        super(executorService);
    }

    public byte[] getContent(InputStream is) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len;
        while ((len = is.read(buffer)) != -1) {
            baos.write(buffer, 0, len);
        }
        byte[] bytes = baos.toByteArray();
        baos.close();
        return bytes;
    }
}
