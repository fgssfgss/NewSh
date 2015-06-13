
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * @author Andrew
 */
public class TemplateEngine {

    Order order;
    List<Student> students;

    String templateFileName = "E:\\unn\\sh5.docx";
    String destination = "E:\\unn\\result\\";
    String xmlFileName = "E:\\unn\\file.xml";
    String docFileName = "E:\\unn\\doc.docx";

    public void parseFiles() {
        XmlParser xParser = new XmlParser(xmlFileName);

        order = xParser.getOrder();
        students = xParser.getStudents();
    }

    public void process() throws IOException, InvalidFormatException {

        for (Student stud : students) {
            String resultFileName = destination.concat(stud.firstName.en.concat(stud.lastName.en.concat(".docx")));

            List<pizdec> r = new ArrayList<>();
            r.add(new pizdec("xml_lname_eng", stud.lastName.en));
            r.add(new pizdec("xml_lname", stud.lastName.ua));

            XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFileName));
            replaceInParagraphs(r, doc.getParagraphs());

            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        replaceInParagraphs(r, cell.getParagraphs());
                    }
                }
            }

            doc.write(new FileOutputStream(resultFileName));

//            XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFileName));
//            for (XWPFParagraph p : doc.getParagraphs()) {
//                List<XWPFRun> runs = p.getRuns();
//                if (runs != null) {
//                    for (XWPFRun r : runs) {
//                        String text = r.getText(0);
//
//                        if (text != null) {
//                            //System.out.println(text);
//                            //System.out.println();
//
//                            //r.setText(text, 0);
//                        }
//                    }
//                }
//            }
//            for (XWPFTable tbl : doc.getTables()) {
//                for (XWPFTableRow row : tbl.getRows()) {
//                    for (XWPFTableCell cell : row.getTableCells()) {
//                        for (XWPFParagraph p : cell.getParagraphs()) {
//                            for (XWPFRun r : p.getRuns()) {
//                                String text = r.getText(0);
//
//                                if (text != null) {
//
//                                    if (text.contains("xml_lname_eng")) {
//                                        text = text.replaceAll("xml_lname_eng", stud.lastName.en);
//                                        System.out.println(text);
//                                        r.setText(text);
//                                    }
//
//                                    if (text.contains("xml_lname")) {
//                                        text = text.replaceAll("xml_lname", stud.lastName.ua);
//                                        r.setText(text);
//                                    }
//
//                                    if (text.contains("xml_fmname_eng")) {
//                                        text = text.replaceAll("xml_fmname_eng", stud.firstName.en);
//                                        r.setText(text);
//                                    }
//
//                                    if (text.contains("xml_fmname")) {
//                                        text = text.replaceAll("xml_fmname", stud.firstName.ua);
//                                        r.setText(text);
//                                    }
//
//                                    //System.out.println(text);
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//            doc.write(new FileOutputStream(resultFileName));
        }
    }

    public long replaceInParagraphs(List<pizdec> replacements, List<XWPFParagraph> xwpfParagraphs) {
        long count = 0;
        for (XWPFParagraph paragraph : xwpfParagraphs) {
            List<XWPFRun> runs = paragraph.getRuns();

            for (pizdec c : replacements) {
                String find = c.src;
                String repl = c.dest;
                TextSegement found = paragraph.searchText(find, new PositionInParagraph());
                if (found != null) {
                    count++;
                    if (found.getBeginRun() == found.getEndRun()) {

                        XWPFRun run = runs.get(found.getBeginRun());
                        String runText = run.getText(run.getTextPosition());
                        String replaced = runText.replace(find, repl);
                        run.setText(replaced, 0);
                    } else {
                        StringBuilder b = new StringBuilder();
                        for (int runPos = found.getBeginRun(); runPos <= found.getEndRun(); runPos++) {
                            XWPFRun run = runs.get(runPos);
                            b.append(run.getText(run.getTextPosition()));
                        }
                        String connectedRuns = b.toString();
                        String replaced = connectedRuns.replace(find, repl);

                        XWPFRun partOne = runs.get(found.getBeginRun());
                        partOne.setText(replaced, 0);

                        for (int runPos = found.getBeginRun() + 1; runPos <= found.getEndRun(); runPos++) {
                            XWPFRun partNext = runs.get(runPos);
                            partNext.setText("", 0);
                        }
                    }
                }
            }
        }
        return count;
    }
}

//            for (Map.Entry<String, String> replPair : replacements.entrySet()) {
//                String find = replPair.getKey();
//                String repl = replPair.getValue();
//                TextSegement found = paragraph.searchText(find, new PositionInParagraph());
//                if (found != null) {
//                    count++;
//                    if (found.getBeginRun() == found.getEndRun()) {
//
//                        XWPFRun run = runs.get(found.getBeginRun());
//                        String runText = run.getText(run.getTextPosition());
//                        String replaced = runText.replace(find, repl);
//                        run.setText(replaced, 0);
//                    } else {
//                        StringBuilder b = new StringBuilder();
//                        for (int runPos = found.getBeginRun(); runPos <= found.getEndRun(); runPos++) {
//                            XWPFRun run = runs.get(runPos);
//                            b.append(run.getText(run.getTextPosition()));
//                        }
//                        String connectedRuns = b.toString();
//                        String replaced = connectedRuns.replace(find, repl);
//
//                        XWPFRun partOne = runs.get(found.getBeginRun());
//                        partOne.setText(replaced, 0);
//
//                        for (int runPos = found.getBeginRun() + 1; runPos <= found.getEndRun(); runPos++) {
//                            XWPFRun partNext = runs.get(runPos);
//                            partNext.setText("", 0);
//                        }
//                    }
//                }
//            }
class pizdec {

    String src;
    String dest;

    public pizdec(String src, String dest) {
        this.src = src;
        this.dest = dest;
    }
};
