package com.github.linyuzai.jpoi.exception;

public class JPoiException extends RuntimeException {
    public JPoiException() {
    }

    public JPoiException(String message) {
        super(message);
    }

    public JPoiException(String message, Throwable cause) {
        super(message, cause);
    }

    public JPoiException(Throwable cause) {
        super(cause);
    }

    public JPoiException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
