package com.wyz.common.commons.util.thread;

import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * @Description: 全局线程池
 * @Author: WangYouzheng
 * @Date: 2023/1/29 13:57
 * @Version: V1.0
 */
public class ThreadPoolUtil {
    public static void main(String[] args) {

        ThreadPoolExecutor fastTriggerPool = new ThreadPoolExecutor(
                10,
                10,
                60L,
                TimeUnit.SECONDS,
                new LinkedBlockingQueue<Runnable>(1000),
                r -> new Thread(r, "wyz-commons, threadPoolHelper-pool-%d" + r.hashCode()));
    }

}
