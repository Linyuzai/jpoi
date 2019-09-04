package com.github.linyuzai.jpoi.order;

public interface Ordered {

    default int getOrder() {
        return 0;
    }
}
