
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
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
    DocParser.DocParsed dp;

    //final String PATH = "D:\\Programming\\NewSh\\doc\\";
    final String PATH = "E:\\unn\\";

    String templateFileName = PATH + "sh9.docx";
    String destination = PATH + "result\\";
    String xmlFileName = PATH + "file2.xml";
    String docFileName = PATH + "file3.docx";

    public void parseFiles() {
        XmlParser xParser = new XmlParser(xmlFileName);
        DocParser docParser = new DocParser(docFileName);
        //docParser.removePassword(); // not work, need to hack
        dp = docParser.parse();
        order = xParser.getOrder();
        students = xParser.getStudents();
    }

    public void process() throws IOException, InvalidFormatException {

        for (Student stud : students) {
            String resultFileName = destination.concat(stud.firstName.en.concat(stud.lastName.en.concat(".docx")));

            List<ReplaceRule> RuleListTable = new ArrayList<>();

            RuleListTable.add(new ReplaceRule("%xml_form%", order.timeEducation));

            if (order.timeEducation.equals("Денна")) {
                RuleListTable.add(new ReplaceRule("%xml_form_eng%", "Full time"));
            }

            RuleListTable.add(new ReplaceRule("%xml_lname_eng%", stud.lastName.en));
            RuleListTable.add(new ReplaceRule("%xml_lname%", stud.lastName.ua));
            RuleListTable.add(new ReplaceRule("%xml_fmname_eng%", stud.firstName.en));
            RuleListTable.add(new ReplaceRule("%xml_fmname%", stud.firstName.ua));

            RuleListTable.add(new ReplaceRule("%xml_honor_eng%", stud.honorEn));
            RuleListTable.add(new ReplaceRule("%xml_honor%", stud.honor));

            RuleListTable.add(new ReplaceRule("%xml_birthday%", stud.birthday));
            RuleListTable.add(new ReplaceRule("%xml_S%", stud.diplom.seria));
            RuleListTable.add(new ReplaceRule("%xml_N%", stud.diplom.number));

            RuleListTable.add(new ReplaceRule("%doc_text21%", dp.tp.p21));
            RuleListTable.add(new ReplaceRule("%doc_text21_eng%", dp.tp.p21_eng));
            RuleListTable.add(new ReplaceRule("%doc_text22%", dp.tp.p22));
            RuleListTable.add(new ReplaceRule("%doc_text22_eng%", dp.tp.p22_eng));
            RuleListTable.add(new ReplaceRule("%doc_text31%", dp.tp.p31));
            RuleListTable.add(new ReplaceRule("%doc_text31_eng%", dp.tp.p31_eng));
            RuleListTable.add(new ReplaceRule("%doc_text32%", dp.tp.p32));
            RuleListTable.add(new ReplaceRule("%doc_text32_eng%", dp.tp.p32_eng));

            RuleListTable.add(new ReplaceRule("%doc_text42_1%", dp.tp.p42_1));
            RuleListTable.add(new ReplaceRule("%doc_text42_1_eng%", dp.tp.p42_1_eng));
            RuleListTable.add(new ReplaceRule("%doc_text42_2%", dp.tp.p42_2));
            RuleListTable.add(new ReplaceRule("%doc_text42_2_eng%", dp.tp.p42_2_eng));
            RuleListTable.add(new ReplaceRule("%doc_text42_3%", dp.tp.p42_3));
            RuleListTable.add(new ReplaceRule("%doc_text42_3_eng%", dp.tp.p42_3_eng));

            Integer marksCount = 0;
            boolean summarySetted = false;
            double det = 0.;
            String lastProgUnit = "";

            for (DocParser.MarkParsed mr : dp.mp) {
                marksCount++;
                for (Discipline d : stud.marks) {
                    if (d.name.equals(mr.subject)) {
                        if (!lastProgUnit.equals(d.programUnit)) {
                            RuleListTable.add(new ReplaceRule("%num".concat(marksCount.toString()).concat("%"), marksCount.toString()));
                            RuleListTable.add(new ReplaceRule("%doc_course".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%doc_year".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%de".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%m".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%xml_grade".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%l".concat(marksCount.toString()).concat("%"), " "));
                            marksCount++;
                        }

                        RuleListTable.add(new ReplaceRule("%num".concat(marksCount.toString()).concat("%"), marksCount.toString()));
                        RuleListTable.add(new ReplaceRule("%doc_course".concat(marksCount.toString()).concat("%"), mr.subject.concat("/").concat(mr.subject_eng)));
                        RuleListTable.add(new ReplaceRule("%doc_year".concat(marksCount.toString()).concat("%"), Integer.toString(mr.year)));
                        RuleListTable.add(new ReplaceRule("%m".concat(marksCount.toString()).concat("%"), d.mark));
                        String de = Float.toString(mr.credits).concat(" (").concat(Integer.toString(mr.hours)).concat(")");
                        RuleListTable.add(new ReplaceRule("%de".concat(marksCount.toString()).concat("%"), de));
                        det += mr.credits;
                        String nMark = "";
                        switch (d.nationalMark) {
                            case "Добре":
                                nMark = "Добре / Good";
                                break;
                            case "Відмінно":
                                nMark = "Відмінно / Excellent";
                                break;
                            case "Задовільно":
                                nMark = "Задовільно / Satisfactory";
                                break;
                            case "Незадовільно":
                                nMark = "Незадовільно / Fail";
                                break;
                            case "Зараховано":
                                nMark = "Зараховано / Passed";
                                break;
                            case "Не зараховано":
                                nMark = "Не зараховано / Fail";
                                break;
                        }

                        RuleListTable.add(new ReplaceRule("%xml_grade".concat(marksCount.toString()).concat("%"), nMark));
                        RuleListTable.add(new ReplaceRule("%l".concat(marksCount.toString()).concat("%"), d.ectsMark));
                        lastProgUnit = d.programUnit;
                        break;
                    } else if (d.programUnit.equals("50") && !summarySetted) {
                        String Mark = "";
                        switch (d.nationalMark) {
                            case "Добре":
                                Mark = "Добре / Good";
                                break;
                            case "Відмінно":
                                Mark = "Відмінно / Excellent";
                                break;
                            case "Задовільно":
                                Mark = "Задовільно / Satisfactory";
                                break;
                            case "Незадовільно":
                                Mark = "Незадовільно / Fail";
                                break;
                        }

                        RuleListTable.add(new ReplaceRule("%mt%", d.mark));
                        RuleListTable.add(new ReplaceRule("%lt%", d.ectsMark));
                        RuleListTable.add(new ReplaceRule("%xml_gradet%", Mark));
                        summarySetted = true;
                    }
                }
            }

            RuleListTable.add(new ReplaceRule("%det%", Double.toString(det)));

            for (Integer i = marksCount; i <= 80; i++) {
                RuleListTable.add(new ReplaceRule("%num".concat(i.toString()).concat("%"), " "));
                RuleListTable.add(new ReplaceRule("%doc_course".concat(i.toString()).concat("%"), " "));
                RuleListTable.add(new ReplaceRule("%doc_year".concat(i.toString()).concat("%"), " "));
                RuleListTable.add(new ReplaceRule("%de".concat(i.toString()).concat("%"), " "));
                RuleListTable.add(new ReplaceRule("%m".concat(i.toString()).concat("%"), " "));
                RuleListTable.add(new ReplaceRule("%xml_grade".concat(i.toString()).concat("%"), " "));
                RuleListTable.add(new ReplaceRule("%l".concat(i.toString()).concat("%"), " "));
            }

            XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFileName));

            replaceInParagraphs(RuleListTable, doc.getParagraphs());

            for (XWPFFooter p : doc.getFooterList()) {
                replaceInParagraphs(RuleListTable, p.getParagraphs());
            }

            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        replaceInParagraphs(RuleListTable, cell.getParagraphs());
                    }
                }
            }

            doc.write(new FileOutputStream(resultFileName));

        }
    }

    public long replaceInParagraphs(List<ReplaceRule> replacements, List<XWPFParagraph> xwpfParagraphs) {
        long count = 0;
        for (XWPFParagraph paragraph : xwpfParagraphs) {
            List<XWPFRun> runs = paragraph.getRuns();

            for (ReplaceRule c : replacements) {
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

class ReplaceRule {

    String src;
    String dest;

    public ReplaceRule(String src, String dest) {
        this.src = src;
        this.dest = dest;
    }
};
