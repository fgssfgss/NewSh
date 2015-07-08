package com.newsh;


import java.io.FileInputStream;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author b2soft
 */
public class DocParser {
    
    String parsedData;
    String fileName;
    XWPFDocument doc;
    private int end;

    public DocParser(String fileName) {
        MyLogger.log("DocParser: constructor");
        this.fileName = fileName;
        try {
            FileInputStream is = new FileInputStream(fileName);
            doc = new XWPFDocument(is);
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    class DocParsed {

        TextParsed tp;
        List<MarkParsed> mp;

        public DocParsed(TextParsed tp, List<MarkParsed> mp) {
            this.tp = tp;
            this.mp = mp;
        }
    }

    public DocParsed parse() {
        MyLogger.log("DocParseD: parse start");
        DocParsed dp;
        XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
        parsedData = extractor.getText().replaceAll("(\\r|\\n)", "");

        MyLogger.log("DocParseD: replace all \\r \\n");

        MyLogger.log("Parsing DOC text...");
        TextParsed tp = parseText();
        MyLogger.log("Parsing DOC complete");

        parsedData = parsedData.substring(end);

        List<MarkParsed> mp = null;

        try {
            MyLogger.log("Parsing marks...");
            mp = parseMarks();
            MyLogger.log("Parsing marks complete");
        } catch (ParseException e) {
            MyLogger.log("EXPT in marks parsing: " + e.toString());
        }
        
        return new DocParsed(tp, mp);
    }

    private TextParsed parseText() {
        PrintWriter writer = null;
        try {
            writer = new PrintWriter("out.txt", "UTF-8");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }

        TextParsed dp = new TextParsed();
        writer.println(parsedData);
        dp.p21 = parseRegEx("\\[2\\.1\\](.*?)3\\.");
        MyLogger.log("2.1 complete");
        writer.println(dp.p21);
        dp.p21_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[2\\.1\\](.*?)4\\.");
        writer.println(dp.p21_eng);
        MyLogger.log("2.1 eng complete");

        dp.p21_2 = parseRegEx("\u0441\u044f\\)\\s\\[2\\.1\\](.*?)5\\.");
        MyLogger.log("2.1_2 complete");
        writer.println(dp.p21_2);
        dp.p21_2_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b.\\)\\s\\(\u044f\u043a\u0449\u043e\\s\u043d\u0430\u0434\u0430\u0454\u0442\u044c\u0441\u044f\\)\\s\\[2\\.1\\](.*?)6\\.");
        MyLogger.log("2.1_2 eng complete");
        writer.println(dp.p21_2_eng);

        dp.p22 = parseRegEx("\\[2\\.2\\](.*?)7\\.");
        MyLogger.log("2.2 complete");
        writer.println(dp.p22);
        dp.p22_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[2\\.2\\](.*?)8\\.");
        MyLogger.log("2.2 eng complete");
        writer.println(dp.p22_eng);

        dp.p31 = parseRegEx("\\[3\\.1\\](.*?)9\\.");
        MyLogger.log("3.1 complete");
        writer.println(dp.p31);
        dp.p31_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[3\\.1\\](.*?)10\\.");
        MyLogger.log("3.1 eng complete");
        writer.println(dp.p31_eng);

        dp.p32 = parseRegEx("\\[3\\.2\\](.*?)11\\.");
        MyLogger.log("3.2 complete");
        writer.println(dp.p32);
        dp.p32_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[3\\.2\\](.*?)12\\.");
        MyLogger.log("3.2 eng complete");
        writer.println(dp.p32_eng);

        dp.p33 = parseRegEx("\\[3\\.3\\](.*?)13\\.");
        MyLogger.log("3.3 complete");
        writer.println(dp.p33);
        dp.p33_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[3\\.3\\](.*?)14\\.");
        MyLogger.log("3.3 eng complete");
        writer.println(dp.p33_eng);

        dp.p42_1 = parseRegEx("\\[4\\.2\\](.*?)15\\.");
        MyLogger.log("4.2_1 complete");
        writer.println(dp.p42_1);
        dp.p42_1_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)16\\.");
        MyLogger.log("4.2_1 eng complete");
        writer.println(dp.p42_1_eng);

        dp.p42_2 = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\s\\[4\\.2\\](.*?)17\\.");
        MyLogger.log("4.2_2 complete");
        writer.println(dp.p42_2);
        dp.p42_2_eng = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)18\\.");
        MyLogger.log("4.2_2 eng complete");
        writer.println(dp.p42_2_eng);

        dp.p42_3 = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\:\\s\\[4\\.2\\](.*?)19\\.");
        MyLogger.log("4.2_3 complete");
        writer.println(dp.p42_3);
        dp.p42_3_eng = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\:\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)20\\.");
        MyLogger.log("4.2_3 eng complete");
        writer.println(dp.p42_3_eng);

        dp.p42_4 = parseRegEx("\u0441\u0443\u0434\u0436\u0435\u043d\u044c\\:\\s\\[4\\.2\\](.*?)21\\.");
        MyLogger.log("4.2_4 complete");
        writer.println(dp.p42_4);
        dp.p42_4_eng = parseRegEx("\u0441\u0443\u0434\u0436\u0435\u043d\u044c\\:\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)22\\.");
        MyLogger.log("4.2_4 eng complete");
        writer.println(dp.p42_4_eng);

        int end = this.end;

        dp.p51 = parseRegEx("\\[5\\.1\\](.*?)25\\.");
        MyLogger.log("5.1 complete");
        writer.println(dp.p51);
        dp.p51_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[5\\.1\\](.*?)26\\.");
        MyLogger.log("5.1 eng complete");
        writer.println(dp.p51_eng);

        dp.p52 = parseRegEx("\\[5\\.2\\](.*?)27\\.");
        MyLogger.log("5.2 complete");
        writer.println(dp.p52);
        dp.p52_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[5\\.2\\](.*?)28\\.");
        MyLogger.log("5.2 eng complete");
        writer.println(dp.p52_eng);

        dp.p65 = parseRegEx("\\[6\\.1\\](.*?)29\\.");
        MyLogger.log("6.5 complete");
        writer.println(dp.p65);
        dp.p65_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[6\\.1\\](.*?)30\\.");
        MyLogger.log("6.5 eng complete");
        writer.println(dp.p65_eng);

        dp.p65_spec = parseRegEx("\u0446\u0456\u044f\\s\\[6\\.1\\](.*?)31\\.");
        MyLogger.log("6.5_spec complete");
        writer.println(dp.p65_spec);
        dp.p65_spec_eng = parseRegEx("\u0446\u0456\u044f\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[6\\.1\\](.*?)32\\.");
        MyLogger.log("6.5_spec eng complete");
        writer.println(dp.p65_spec_eng);

        this.end = end;

        writer.close();

        return dp;

    }

    private List<MarkParsed> parseMarks() throws ParseException {
        List<MarkParsed> parsedList = new ArrayList<>();
        MyLogger.log("Preparing table");
        String tableData = parsedData;
        Pattern pattern = Pattern.compile("\\*\\s\u00ab\u041d\u043e\u043c\u0435\u0440\\s\u0437\u0430\\s\u043f\u043e\u0440\u044f\u0434\u043a\u043e\u043c");
        Matcher matcher = pattern.matcher(tableData);
        if (matcher.find()) {
            tableData = parsedData.substring(0, matcher.end());
        }
        MyLogger.log(String.format("tableData %s", tableData));    
        pattern = Pattern.compile("\\t(201\\d/201\\d\\t.*?\\t.*?\\t.*?\\t\\d{2,3}\\t\\d{1,2})");
        matcher = pattern.matcher(tableData);
        List<String> table = new ArrayList<>();
        while (matcher.find()) {
            MyLogger.log(String.format("matcher %s", matcher.group(1)));
            table.add(matcher.group(1));
        }
        MyLogger.log("All marks in list now. Starting parse each");

        int currType = 1;
        int i = 1;
        for (String e : table) {
            MyLogger.log(String.format("Parsing %d/%d", i++, table.size()));
            String[] sparse = e.split("\\t");
            String year = sparse[0];
            String subject = sparse[1];
            String subject_eng = sparse[2];

            NumberFormat format = NumberFormat.getInstance(Locale.FRANCE);
            Number number = format.parse(sparse[3]);
            float credits = number.floatValue();

            int hours = Integer.valueOf(sparse[4]);
            int markType = Integer.valueOf(sparse[5].substring(0, 1));

            parsedList.add(new MarkParsed((sparse[5].length() == 2) ? currType++ : currType, year, subject, subject_eng, credits, hours, markType));

        }
        return parsedList;
    }

    private String parseRegEx(String regex) {
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(parsedData);
        if (matcher.find()) {
            end = matcher.end();
            return matcher.group(1);
        }
        return "NOT FOUND";
    }

    public static class MarkParsed {

        int type; // 1 - ?????????? ???????? 2 - ???????? 3 - ??????? ?????? (???????) 4 - ?????????? ???????? ?????????
        String year;
        String subject;
        String subject_eng;
        float credits;
        int hours;
        int mark_type; //0 - ekz, 1 - zalik

        public MarkParsed(int type, String year, String subject, String subject_eng, float credits, int hours, int mark_type) {
            this.type = type;
            this.year = year;
            this.subject = subject;
            this.subject_eng = subject_eng;
            this.credits = credits;
            this.hours = hours;
            this.mark_type = mark_type;
            
            System.out.println(toString());
        }

        @Override
        public String toString() {
            return "MarkParsed{"
                    + "type=" + type
                    + ", year=" + year
                    + ", subject='" + subject + '\''
                    + ", subject_eng='" + subject_eng + '\''
                    + ", credits=" + credits
                    + ", hours=" + hours
                    + ", mark_type=" + mark_type
                    + '}';
        }
    }

    public static class TextParsed {

        String p21;
        String p21_eng;
        String p21_2;
        String p21_2_eng;
        String p22;
        String p22_eng;
        String p31;
        String p31_eng;
        String p32;
        String p32_eng;
        String p33;
        String p33_eng;
        String p42_1;
        String p42_1_eng;
        String p42_2;
        String p42_2_eng;
        String p42_3;
        String p42_3_eng;
        String p42_4;
        String p42_4_eng;
        String p51;
        String p51_eng;
        String p52;
        String p52_eng;
        String p65;
        String p65_eng;
        String p65_spec;
        String p65_spec_eng;
    }
}
