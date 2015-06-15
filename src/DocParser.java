import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

/**
 * @author Andrew
 */

public class DocParser {

    String parsedData;
    String fileName;
    XWPFDocument doc;

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


    public void parse() {
        XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
        PrintWriter writer = null;
        try {
            writer = new PrintWriter("out.txt", "UTF-8");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }

        writer.println(extractor.getText());
        writer.close();
    }
}
