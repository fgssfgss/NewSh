package com.newsh;

import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * Created by Bogdan on 7/6/2015.
 */
public class MyLogger {

    static {
        debugMode = false;
        try {
            logFile =  new PrintWriter("log.txt", "UTF-8");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
    }
    private static boolean debugMode;
    private static PrintWriter logFile;


    public static void setDebugMode() { debugMode = true; }
    public static void setNormalMode() { debugMode = false; }
    public static void saveLog() { if (logFile != null) logFile.close(); }

    public synchronized static void log(String text) {
        Date dNow = new Date();

        SimpleDateFormat ft = new SimpleDateFormat ("[HH:mm:ss:SSS]");
        logFile.println(ft.format(dNow) + " " + text);
       // System.out.println(ft.format(dNow) + " " + text);
        logFile.flush();
    }



}
