import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * @author Andrew
 */
public class XmlParser {

    private Element root;
    private NodeList studentList;

    public XmlParser(String filename) {
        try {
            File file = new File(filename);
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.parse(file);
            root = document.getDocumentElement();
            studentList = root.getElementsByTagName("student");

        } catch (ParserConfigurationException | SAXException | IOException ex) {
        }
    }

    public Order getOrder() {
        Order ord = new Order();

        Element elem = (Element) ((Element) root.getElementsByTagName("order").item(0));

        ord.timeEducation = elem.getElementsByTagName("time_education").item(0).getTextContent();
        ord.issued = elem.getElementsByTagName("issued").item(0).getTextContent();
        ord.graduated = elem.getElementsByTagName("graduated").item(0).getTextContent();
        ord.faculty = elem.getElementsByTagName("faculty").item(0).getAttributes().getNamedItem("uk").getTextContent();
        ord.qualification = elem.getElementsByTagName("qualification").item(0).getAttributes().getNamedItem("uk").getTextContent();

        return ord;
    }

    public List<Student> getStudents() {
        List<Student> students = new ArrayList<>();

        for (int index = 0; index < studentList.getLength(); index++) {
            Element element = (Element) ((Element) studentList.item(index));
            Student student = new Student();

            student.firstName.en = element.getElementsByTagName("first_name").item(0).getAttributes().getNamedItem("en").getTextContent();
            student.firstName.ua = element.getElementsByTagName("first_name").item(0).getAttributes().getNamedItem("uk").getTextContent();

            student.lastName.en = element.getElementsByTagName("last_name").item(0).getAttributes().getNamedItem("en").getTextContent();
            student.lastName.ua = element.getElementsByTagName("last_name").item(0).getAttributes().getNamedItem("uk").getTextContent();

            student.middleName = element.getElementsByTagName("middle_name").item(0).getAttributes().getNamedItem("uk").getTextContent();

            student.sex = element.getElementsByTagName("sex").item(0).getTextContent();

            student.birthday = element.getElementsByTagName("birthday").item(0).getTextContent();

            student.personDocument.id = element.getElementsByTagName("person_document").item(0).getAttributes().getNamedItem("ID").getTextContent();
            student.personDocument.number = element.getElementsByTagName("person_document").item(0).getAttributes().getNamedItem("number").getTextContent();
            student.personDocument.seria = element.getElementsByTagName("person_document").item(0).getAttributes().getNamedItem("seria").getTextContent();

            student.honor = element.getElementsByTagName("honor").item(0).getTextContent();

            student.prevDocument.id = element.getElementsByTagName("prev_document").item(0).getAttributes().getNamedItem("ID").getTextContent();
            student.prevDocument.number = element.getElementsByTagName("prev_document").item(0).getAttributes().getNamedItem("number").getTextContent();
            student.prevDocument.seria = element.getElementsByTagName("prev_document").item(0).getAttributes().getNamedItem("seria").getTextContent();

            student.prevQualification.en = element.getElementsByTagName("prev_qualification").item(0).getAttributes().getNamedItem("en").getTextContent();
            student.prevQualification.ua = element.getElementsByTagName("prev_qualification").item(0).getAttributes().getNamedItem("uk").getTextContent();

            student.prevSpeciality.en = element.getElementsByTagName("prev_speciality").item(0).getAttributes().getNamedItem("en").getTextContent();
            student.prevSpeciality.ua = element.getElementsByTagName("prev_speciality").item(0).getAttributes().getNamedItem("uk").getTextContent();

            student.receiptDay = element.getElementsByTagName("receipt_date").item(0).getTextContent();

            student.payment = element.getElementsByTagName("payment").item(0).getTextContent();

            NodeList marks = ((Element)element.getElementsByTagName("marks").item(0)).getElementsByTagName("discipline");

            for (int mindex = 0; mindex < marks.getLength(); mindex++) {
                Element mark = (Element)(marks.item(mindex));
                Discipline di = new Discipline();
                
                di.name = mark.getAttribute("name");
                di.mark = mark.getAttribute("mark");
                di.nationalMark = mark.getAttribute("national_mark");
                di.ectsMark = mark.getAttribute("ects_mark");
                di.programUnit = mark.getAttribute("program_unit");
                
                student.marks.add(di);
            }

            students.add(student);
        }

        return students;
    }

}
