
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Locale;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFSettings;

/**
 * @author b2soft
 */
public class DocParser {

    String parsedData;
    String fileName;
    XWPFDocument doc;
    private int end;

    public DocParser(String fileName) {
        this.fileName = fileName;
        try {
            doc = new XWPFDocument(OPCPackage.open(fileName));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    static String convertStreamToString(java.io.InputStream is) {
        java.util.Scanner s = new java.util.Scanner(is).useDelimiter("\\A");
        return s.hasNext() ? s.next() : "";
    }

    public void removePassword() {
        doc.removeProtectionEnforcement();
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
        DocParsed dp;
        XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
        parsedData = extractor.getText().replaceAll("(\\r|\\n)", "");

        System.out.println(parsedData);
        
        TextParsed tp = parseText();

        parsedData = parsedData.substring(end);
        System.out.println(end);

        List<MarkParsed> mp = null;

        try {
            mp = parseMarks();
        } catch (ParseException e) {
            e.printStackTrace();
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
        dp.p21 = parseRegEx("\\[2\\.1\\](.*?)3");
        writer.println(dp.p21);
        System.out.println(dp.p21);
        dp.p21_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[2\\.1\\](.*?)4");
        writer.println(dp.p21_eng);
        System.out.println(dp.p21_eng);

        dp.p22 = parseRegEx("\\[2\\.2\\](.*?)7");
        writer.println(dp.p22);
        dp.p22_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[2\\.2\\](.*?)8");
        writer.println(dp.p22_eng);

        dp.p31 = parseRegEx("\\[3\\.1\\](.*?)9");
        writer.println(dp.p31);
        dp.p31_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[3\\.1\\](.*?)10");
        writer.println(dp.p31_eng);

        dp.p32 = parseRegEx("\\[3\\.2\\](.*?)11");
        writer.println(dp.p32);
        dp.p32_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[3\\.2\\](.*?)12");
        writer.println(dp.p32_eng);

        dp.p33 = parseRegEx("\\[3\\.3\\](.*?)13");
        writer.println(dp.p33);
        dp.p33_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[3\\.3\\](.*?)14");
        writer.println(dp.p33_eng);

        dp.p42_1 = parseRegEx("\\[4\\.2\\](.*?)15");
        writer.println(dp.p42_1);
        dp.p42_1_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)16");
        writer.println(dp.p42_1_eng);

        dp.p42_2 = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\s\\[4\\.2\\](.*?)17");
        writer.println(dp.p42_2);
        dp.p42_2_eng = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)18");
        writer.println(dp.p42_2_eng);

        dp.p42_3 = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\:\\s\\[4\\.2\\](.*?)19");
        writer.println(dp.p42_3);
        dp.p42_3_eng = parseRegEx("\u0440\u043e\u0437\u0443\u043c\u0456\u043d\u043d\u044f\\:\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)20");
        writer.println(dp.p42_3_eng);

        dp.p42_4 = parseRegEx("\u0441\u0443\u0434\u0436\u0435\u043d\u044c\\:\\s\\[4\\.2\\](.*?)21");
        writer.println(dp.p42_4);
        dp.p42_4_eng = parseRegEx("\u0441\u0443\u0434\u0436\u0435\u043d\u044c\\:\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[4\\.2\\](.*?)22");
        writer.println(dp.p42_4_eng);

        int end = this.end;

        dp.p51 = parseRegEx("\\[5\\.1\\](.*?)25");
        writer.println(dp.p51);
        dp.p51_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[5\\.1\\](.*?)26");
        writer.println(dp.p51_eng);

        dp.p52 = parseRegEx("\\[5\\.2\\](.*?)27");
        writer.println(dp.p52);
        dp.p52_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[5\\.2\\](.*?)28");
        writer.println(dp.p52_eng);

        dp.p65 = parseRegEx("\\[6\\.1\\](.*?)29");
        writer.println(dp.p65);
        dp.p65_eng = parseRegEx("\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[6\\.1\\](.*?)30");
        writer.println(dp.p65_eng);

        dp.p65_spec = parseRegEx("\u0446\u0456\u044f\\s\\[6\\.1\\](.*?)31");
        writer.println(dp.p65_spec);
        dp.p65_spec_eng = parseRegEx("\u0446\u0456\u044f\\s\\(\u0430\u043d\u0433\u043b\\.\\)\\s\\[6\\.1\\](.*?)32");
        writer.println(dp.p65_spec_eng);

        this.end = end;

        writer.close();
        return dp;

    }

    private List<MarkParsed> parseMarks() throws ParseException {
        List<MarkParsed> parsedList = new ArrayList<>();

        String tableData = parsedData;
        Pattern pattern = Pattern.compile("\\*\\s\u00ab\u041d\u043e\u043c\u0435\u0440\\s\u0437\u0430\\s\u043f\u043e\u0440\u044f\u0434\u043a\u043e\u043c");
        Matcher matcher = pattern.matcher(tableData);
        if (matcher.find()) {
            System.out.println(end + " " + matcher.end());
            tableData = parsedData.substring(0, matcher.end());
        }

        pattern = Pattern.compile("\\t(201\\d\\t.*?\\t.*?\\t.*?\\t\\d{2,3}\\t\\d{1,2})");
        matcher = pattern.matcher(tableData);
        List<String> table = new ArrayList<>();
        while (matcher.find()) {
            table.add(matcher.group(1));
        }

        int currType = 1;
        for (String e : table) {
            String[] sparse = e.split("\\t");
            int year = Integer.valueOf(sparse[0]);
            String subject = sparse[1];
            String subject_eng = sparse[2];

            NumberFormat format = NumberFormat.getInstance(Locale.FRANCE);
            Number number = format.parse(sparse[3]);
            float credits = number.floatValue();

            int hours = Integer.valueOf(sparse[4]);
            int markType = Integer.valueOf(sparse[5].substring(0, 1));

            parsedList.add(new MarkParsed((sparse[5].length() == 2) ? currType++ : currType, year, subject, subject_eng, credits, hours, markType));
            System.out.println(parsedList.get(parsedList.size() - 1));

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
        int year;
        String subject;
        String subject_eng;
        float credits;
        int hours;
        int mark_type; //0 - ekz, 1 - zalik

        public MarkParsed(int type, int year, String subject, String subject_eng, float credits, int hours, int mark_type) {
            this.type = type;
            this.year = year;
            this.subject = subject;
            this.subject_eng = subject_eng;
            this.credits = credits;
            this.hours = hours;
            this.mark_type = mark_type;
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
