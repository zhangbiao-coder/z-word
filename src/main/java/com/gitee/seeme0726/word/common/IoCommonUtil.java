package com.gitee.seeme0726.word.common;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class IoCommonUtil {

    /**
     * 复制InputStream流
     *
     * @param input 输入路
     * @return ByteArrayOutputStream
     * @throws IOException IOException
     */
    public static ByteArrayOutputStream cloneInputStream(InputStream input) throws IOException {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024];
            int len;
            while ((len = input.read(buffer)) > -1) {
                baos.write(buffer, 0, len);
            }
            baos.flush();
            return baos;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

}
