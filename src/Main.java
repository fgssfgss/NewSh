
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Main {

    public static void main(String[] args) {
        try {
            new Main().run();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void run() throws FileNotFoundException, IOException, InvalidFormatException {
        //final String PATH = "D:\\Programming\\NewSh\\doc\\";
        final String PATH = "E:\\unn\\";

        String templateFileName = PATH + "sh9.docx";
        String destination = PATH + "result\\";
        String xmlFileName = PATH + "f1.xml";
        String docFileName = PATH + "file2.docx";

        TemplateEngine tempEngine = new TemplateEngine();
        tempEngine.setDestination(destination);
        tempEngine.setDocFileName(docFileName);
        tempEngine.setTemplateFileName(templateFileName);
        tempEngine.setXmlFileName(xmlFileName);
        tempEngine.parseFiles();
        tempEngine.process();
    }

}
