package app.converter;

import java.util.HashMap;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        Map<String,String> docAtr=new HashMap<String,String>();
        docAtr.put("developer","Сидоров");
        docAtr.put("checker","Иванов");
        docAtr.put("approver","Петров");
        docAtr.put("productName","Изделие А");
        docAtr.put("documentName","Расчет надежности");
        docAtr.put("documentCode","XXXX.231765.879 РР1");
        docAtr.put("documentType","Text");
        app.converter.StyledDocument doc = new app.converter.StyledDocument("D:/base-template.docx",docAtr);
        doc.createFile("D:/StyledDocument.docx");
    }
}