package com.newsh;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Andrew
 */
class Discipline {

    String name;
    String mark;
    String nationalMark;
    String ectsMark;
    String programUnit;
}

public class Student {

    public FirstName firstName;
    public LastName lastName;
    public PersonDocument personDocument;
    public Diplom diplom;
    public PrevDocument prevDocument;
    public PrevQualification prevQualification;
    public PrevSpeciality prevSpeciality;
    public String middleName;
    public String sex;
    public String h;
    public String birthday;
    public String honor;
    public String honorEn;
    public String receiptDay;
    public String payment;
    public List<Discipline> marks;

    public Student() {
        marks = new ArrayList<>();
        firstName = new FirstName();
        lastName = new LastName();
        personDocument = new PersonDocument();
        diplom = new Diplom();
        prevDocument = new PrevDocument();
        prevQualification = new PrevQualification();
        prevSpeciality = new PrevSpeciality();
    }

    class FirstName {

        String ua;
        String en;
    }

    class LastName {

        String ua;
        String en;
    }

    class PersonDocument {

        String id;
        String number;
        String seria;
    }

    class Diplom {

        String seria;
        String number;
    }

    class PrevDocument {
        String foreigner;
        String id;
        String number;
        String seria;
    }

    class PrevQualification {

        String ua;
        String en;
    }

    class PrevSpeciality {

        String ua;
        String en;
    }

}
