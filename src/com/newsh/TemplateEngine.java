package com.newsh;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
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
public class TemplateEngine extends Thread{
    private final MainFrame frame;
    private boolean st = false;
    private Order order;
    private List<Student> students;
    private DocParser.DocParsed dp;

    private String templateFileName = "";
    private String destination = "";
    private String xmlFileName = "";
    private String docFileName = "";

    public TemplateEngine(MainFrame frame) {
        this.frame = frame;
        st = frame.getStGroupRadioButton().isSelected();
    }
    
    @Override
    public void run(){
        parseFiles();
        try {
            process();
        } catch (IOException | InvalidFormatException ex) {
            Logger.getLogger(TemplateEngine.class.getName()).log(Level.SEVERE, null, ex);
        }

        frame.getMakeItButton().setEnabled(true);
    }
    
    public void setTemplateFileName(String templateFileName) {
        this.templateFileName = templateFileName;
    }

    public void setDestination(String destination) {
        this.destination = destination;
    }

    public void setXmlFileName(String xmlFileName) {
        this.xmlFileName = xmlFileName;
    }

    public void setDocFileName(String docFileName) {
        this.docFileName = docFileName;
    }

    public String EngMark(String nm) {
        String nMark = "";

        switch (nm) {
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

        return nMark;
    }

    public String progUnitName(int num) {
        switch (num) {
            case 1:
                return "Теоретичне навчання / Theoretical classes";
            case 2:
                return "Практики / Practices";
            case 3:
                return "Курсові роботи (проекти) / Term papers (projects)";
            case 4:
                return "Підсумкова державна атестація / Final state certification";
            default:
                return "";
        }
    }

    public void parseFiles() {
        XmlParser xParser = new XmlParser(xmlFileName, 1);
        students = xParser.getStudents();
        order = xParser.getOrder();
        
        DocParser docParser = new DocParser(docFileName);
        dp = docParser.parse();
    }

    public void process() throws IOException, InvalidFormatException {
        frame.updateProgress(0, students.size());
        int progress = 0;
        for (Student stud : students) {
            MyLogger.log(String.format("Creating %d/%d student", progress + 1, students.size()));

            String resultFileName = destination.concat(stud.firstName.en.concat(stud.lastName.en.concat(".docx")));

            List<ReplaceRule> RuleListTable = new ArrayList<>();

            
            String name = frame.getComboBox().getSelectedItem().toString();
            if(stud.h.equals("1")){
                name = name.concat(" з вiдзнакою").toUpperCase();
            } else {
                name = name.toUpperCase();
            }
            
            RuleListTable.add(new ReplaceRule("%Dname%", name));
            RuleListTable.add(new ReplaceRule("%Ddate%", order.graduated));
            RuleListTable.add(new ReplaceRule("%xml_form%", order.timeEducation));

            if (order.timeEducation.equals("Денна")) {
                RuleListTable.add(new ReplaceRule("%xml_form_eng%", "Full time"));
            }

            RuleListTable.add(new ReplaceRule("%xml_lname_eng%", stud.lastName.en));
            RuleListTable.add(new ReplaceRule("%xml_lname%", stud.lastName.ua));
            RuleListTable.add(new ReplaceRule("%xml_fmname_eng%", stud.firstName.en));
            RuleListTable.add(new ReplaceRule("%xml_fmname%", stud.firstName.ua.concat(" ").concat(stud.middleName)));

            RuleListTable.add(new ReplaceRule("%xml_honor_eng%", stud.honorEn));
            RuleListTable.add(new ReplaceRule("%xml_honor%", stud.honor));

            RuleListTable.add(new ReplaceRule("%xml_birthday%", stud.birthday));
            RuleListTable.add(new ReplaceRule("%xml_S%", stud.diplom.seria));
            RuleListTable.add(new ReplaceRule("%xml_N%", stud.diplom.number));

            RuleListTable.add(new ReplaceRule("%doc_text21%", dp.tp.p21));
            RuleListTable.add(new ReplaceRule("%doc_text21_eng%", dp.tp.p21_eng));
            RuleListTable.add(new ReplaceRule("%doc_text21_2%", dp.tp.p21_2));
            RuleListTable.add(new ReplaceRule("%doc_text21_2_eng%", dp.tp.p21_2_eng));
            RuleListTable.add(new ReplaceRule("%doc_text22%", dp.tp.p22));
            RuleListTable.add(new ReplaceRule("%doc_text22_eng%", dp.tp.p22_eng));
            RuleListTable.add(new ReplaceRule("%doc_text31%", dp.tp.p31));
            RuleListTable.add(new ReplaceRule("%doc_text31_eng%", dp.tp.p31_eng));
            RuleListTable.add(new ReplaceRule("%doc_text32%", dp.tp.p32));
            RuleListTable.add(new ReplaceRule("%doc_text32_eng%", dp.tp.p32_eng));
            RuleListTable.add(new ReplaceRule("%doc_text33%", dp.tp.p33));
            RuleListTable.add(new ReplaceRule("%doc_text33_eng%", dp.tp.p33_eng));

            RuleListTable.add(new ReplaceRule("%doc_text42_1%", dp.tp.p42_1));
            RuleListTable.add(new ReplaceRule("%doc_text42_1_eng%", dp.tp.p42_1_eng));
            RuleListTable.add(new ReplaceRule("%doc_text42_2%", dp.tp.p42_2));
            RuleListTable.add(new ReplaceRule("%doc_text42_2_eng%", dp.tp.p42_2_eng));
            RuleListTable.add(new ReplaceRule("%doc_text42_3%", dp.tp.p42_3));
            RuleListTable.add(new ReplaceRule("%doc_text42_3_eng%", dp.tp.p42_3_eng));
            RuleListTable.add(new ReplaceRule("%doc_text42_4%", dp.tp.p42_4));
            RuleListTable.add(new ReplaceRule("%doc_text42_4_eng%", dp.tp.p42_4_eng));

            RuleListTable.add(new ReplaceRule("%doc_text51%", dp.tp.p51));
            RuleListTable.add(new ReplaceRule("%doc_text51_eng%", dp.tp.p51_eng));
            RuleListTable.add(new ReplaceRule("%doc_text52%", dp.tp.p52));
            RuleListTable.add(new ReplaceRule("%doc_text52_eng%", dp.tp.p52_eng));

            String text61 = "";
            String text61Eng = "";
            if (st) {
                text61 = "2 навчальні роки";
                text61Eng = "2 academic years";
            } else {
                text61 = "4 навчальні роки";
                text61Eng = "4 academic years";
            }
            RuleListTable.add(new ReplaceRule("%doc_text61%", text61));
            RuleListTable.add(new ReplaceRule("%doc_text61_eng%", text61Eng));

            String text64 = "Попереднiй документ про освiту / Pregoing document on education: ".concat(stud.prevDocument.seria).concat(" №").concat(stud.prevDocument.number);
            String text64Eng = "-освiтньо-квалiфiкацiйний рiвень / qualification level of education - ";
            switch (stud.prevDocument.id) {
                case "2":
                    text64Eng = text64Eng.concat("Атестат про середню освiту / Atestat");
                    break;
                case "6":
                    text64Eng = text64Eng.concat("Диплом молодшого спецiалiста / Junior Specialist Diploma");
                    break;
            }
            RuleListTable.add(new ReplaceRule("%xml_text64%", text64));
            RuleListTable.add(new ReplaceRule("%xml_text64_eng%", text64Eng));

            RuleListTable.add(new ReplaceRule("%xml_text65_enter%", stud.receiptDay));
            RuleListTable.add(new ReplaceRule("%xml_text65_exit%", order.graduated));
            
            RuleListTable.add(new ReplaceRule("%doc_text65%", dp.tp.p65));
            RuleListTable.add(new ReplaceRule("%doc_text65_eng%", dp.tp.p65_eng));
            RuleListTable.add(new ReplaceRule("%doc_spec_text65%", dp.tp.p65_spec));
            RuleListTable.add(new ReplaceRule("%doc_spec_text65_eng%", dp.tp.p65_spec_eng));
            
            Integer marksCount = 0;
            Integer tablePos = 0;
            Integer progUn = 0;
            boolean summarySetted = false;
            double det = 0.;
            String lastProgUnit = "";

            MyLogger.log(String.format("Marks №%d", dp.mp.size()));
            for (DocParser.MarkParsed mr : dp.mp) {
                marksCount++;
                MyLogger.log(mr.subject);
                for (Discipline d : stud.marks) {
                    MyLogger.log(String.format("Disc №%d:%s", marksCount, mr.subject));
                    
                    if (d.name.equals(mr.subject)) {
                        if (!lastProgUnit.equals(d.programUnit)) {
                            progUn++;
                            RuleListTable.add(new ReplaceRule("%num".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%doc_course".concat(marksCount.toString()).concat("%"), progUnitName(progUn)));
                            RuleListTable.add(new ReplaceRule("%doc_year".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%de".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%m".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%xml_grade".concat(marksCount.toString()).concat("%"), " "));
                            RuleListTable.add(new ReplaceRule("%l".concat(marksCount.toString()).concat("%"), " "));
                            marksCount++;
                        }
                        
                        MyLogger.log(mr.subject);
                        
                        tablePos++;
                        RuleListTable.add(new ReplaceRule("%num".concat(marksCount.toString()).concat("%"), tablePos.toString()));
                        RuleListTable.add(new ReplaceRule("%doc_course".concat(marksCount.toString()).concat("%"), mr.subject.concat("/").concat(mr.subject_eng)));
                        RuleListTable.add(new ReplaceRule("%doc_year".concat(marksCount.toString()).concat("%"), mr.year));
                        RuleListTable.add(new ReplaceRule("%m".concat(marksCount.toString()).concat("%"), d.mark));
                        
                        String de;
                        if (mr.credits - Math.floor(mr.credits) > 0.1) {
                            de = Float.toString(mr.credits).concat(" (").concat(Integer.toString(mr.hours)).concat(")");
                        } else {
                            de = Integer.toString((int) mr.credits).concat(" (").concat(Integer.toString(mr.hours)).concat(")");
                        }
                        RuleListTable.add(new ReplaceRule("%de".concat(marksCount.toString()).concat("%"), de));
                        det += mr.credits;

                        RuleListTable.add(new ReplaceRule("%xml_grade".concat(marksCount.toString()).concat("%"), EngMark(d.nationalMark)));
                        RuleListTable.add(new ReplaceRule("%l".concat(marksCount.toString()).concat("%"), d.ectsMark));
                        lastProgUnit = d.programUnit;
                        break;
                    } else if (d.programUnit.equals("50") && !summarySetted) {
                        RuleListTable.add(new ReplaceRule("%mt%", d.mark));
                        RuleListTable.add(new ReplaceRule("%lt%", d.ectsMark));
                        RuleListTable.add(new ReplaceRule("%xml_gradet%", EngMark(d.nationalMark)));
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

            try (FileOutputStream output = new FileOutputStream(resultFileName)) {
                doc.write(output);
                output.flush();
            }
            progress++;
            MyLogger.log("DONE");
            frame.updateProgress(progress);
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
