package app.servlets;
import app.converter.StyledDocument;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.servlet.http.HttpServlet;
import javax.servlet.ServletException;
import javax.servlet.ServletConfig;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class MainServlet extends HttpServlet {
    public void init(ServletConfig config) throws ServletException {
        super.init();


    }
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws IOException{
        int maxFileSize=20848820;
        DiskFileItemFactory factory = new DiskFileItemFactory(maxFileSize,new File(System.getProperty("java.io.tmpdir")));
        ServletFileUpload upload = new ServletFileUpload(factory);
        upload.setSizeMax(maxFileSize);
        List<FileItem> items;
        try {
            items = upload.parseRequest(request);
        } catch (FileUploadException e) {
            throw new RuntimeException(e);
        }
        Iterator<FileItem> iter = items.iterator();
        Map<String,String> docAtr=new HashMap<String,String>();
        InputStream uploadedStream = null;
        while (iter.hasNext()) {
            FileItem item = iter.next();
            if (item.isFormField()) {
                System.out.println(item.getFieldName());
                System.out.println(new String(item.getString().getBytes(StandardCharsets.ISO_8859_1), StandardCharsets.UTF_8));
                docAtr.put(item.getFieldName(),new String(item.getString().getBytes(StandardCharsets.ISO_8859_1), StandardCharsets.UTF_8));
            } else {
                System.out.println(item.getFieldName());
                System.out.println(item.getName());
                System.out.println(item.getContentType());
                System.out.println("Is in memory"+item.isInMemory());
                System.out.println("Size"+item.getSize());
                uploadedStream = item.getInputStream();
            }

        }
        StyledDocument styledDocument= null;
        try {
            styledDocument = new StyledDocument(uploadedStream,docAtr);
            String fileName="DocWithFrames.docx";
            response.setContentType("application/msword");
            response.setHeader("Content-Disposition","attachment; filename=\""
                    + fileName + "\"");
            OutputStream out = response.getOutputStream();
            styledDocument.createStream(out);
            out.close();
        } catch (InvalidFormatException e) {
            PrintWriter writer=response.getWriter();
            writer.println("InvalidFormatException"+e.toString());
        }
         catch (IOException e) {
             PrintWriter writer=response.getWriter();
             writer.println("Not able to load style template"+e.toString());
        }

    }
}
