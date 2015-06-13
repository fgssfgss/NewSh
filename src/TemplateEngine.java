
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

/**
 * @author Andrew
 */
public class TemplateEngine {

    Order order;
    List<Student> students;

    String templateFileName = "E:\\unn\\sh6.docx";
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

            List<ReplaceRule> RuleListTable = new ArrayList<>();

            RuleListTable.add(new ReplaceRule("xml_lname_eng", stud.lastName.en));
            RuleListTable.add(new ReplaceRule("xml_lname", stud.lastName.ua));
            RuleListTable.add(new ReplaceRule("xml_fmname_eng", stud.firstName.en));
            RuleListTable.add(new ReplaceRule("xml_fmname", stud.firstName.ua));

            RuleListTable.add(new ReplaceRule("xml_birthday", stud.birthday));
            RuleListTable.add(new ReplaceRule("xml_S", stud.diplom.seria));
            RuleListTable.add(new ReplaceRule("xml_N", stud.diplom.number));

            XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFileName));

            replaceInParagraphs(RuleListTable, doc.getParagraphs());

            // footer in pages
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

    private void printContentsOfTextBox(XWPFParagraph paragraph) {

        XmlObject[] textBoxObjects = paragraph.getCTP().selectPath("declare namespace w ='http://schemas.openxmlformats.org/wordprocessingml/2006/main' declare namespace wps ='http://schemas.microsoft.com/office/word/2010/wordprocessingShape' .//*/wps:txbx/w:txbxContent");

        for (int i = 0; i < textBoxObjects.length; i++) {
            XWPFParagraph embeddedPara = null;
            try {
                XmlObject[] paraObjects = textBoxObjects[i].selectChildren(new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p"));

                for (int j = 0; j < paraObjects.length; j++) {
                    embeddedPara = new XWPFParagraph(CTP.Factory.parse(paraObjects[j].xmlText()), paragraph.getBody());
                    System.out.println(embeddedPara.getText());
                    
                }

            } catch (XmlException e) {
            }
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
