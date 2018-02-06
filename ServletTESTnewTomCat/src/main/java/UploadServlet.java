import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.*;

/**
 * Created by Odmin on 14.08.2017.
 */
public class UploadServlet extends HttpServlet {

    String sessionId;
    String home_pth=System.getProperty("user.dir")+"/out";;

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        //проверяем является ли полученный запрос multipart/form-data

        boolean isMultipart = ServletFileUpload.isMultipartContent(request);
        if (!isMultipart) {
            response.sendError(HttpServletResponse.SC_BAD_REQUEST);
            nextPage("","Загруженные файлы",request,response);
            return;
        }
        sessionId = request.getSession().getId();
        // Создаём класс фабрику
        DiskFileItemFactory factory = new DiskFileItemFactory();

        // Максимальный буфера данных в байтах,
        // при его привышении данные начнут записываться на диск во временную директорию
        // устанавливаем один мегабайт
        factory.setSizeThreshold(102400*1024);

        // устанавливаем временную директорию
        File tempDir = (File)getServletContext().getAttribute("javax.servlet.context.tempdir");
        factory.setRepository(tempDir);

        //Создаём сам загрузчик
        ServletFileUpload upload = new ServletFileUpload(factory);

        //максимальный размер данных который разрешено загружать в байтах
        //по умолчанию -1, без ограничений. Устанавливаем 100 мегабайт.
        //    upload.setSizeMax(1024 * 1024 * 100);

        try {
            List items = upload.parseRequest(request);
            Iterator iter = items.iterator();
            while (iter.hasNext()) {
                FileItem item = (FileItem) iter.next();

                if (item.isFormField()) {
                    //если принимаемая часть данных является полем формы
                    processFormField(item);
                    nextPage("Что-то не так с файлом","Загруженные файлы",request,response);
                } else {
                    //в противном случае рассматриваем как файл
                    if(processUploadedFile(item, sessionId ))
                    {
                        File folder = new File(home_pth+"//"+ sessionId);
                        nextPage("Файл успешно загружен","Загруженные файлы",request,response);
                    }
                    else
                        nextPage("Что-то не так с файлом","Загруженные файлы",request,response);

                }

            }
        } catch (Exception e) {
            e.printStackTrace();
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);

            nextPage("Что-то не так с файлом","Загруженные файлы",request,response);

            return;
        }

        // response.getWriter().println("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"></head><body>SUCCESS!<body><p><a href=\"index.html\">Back</a></p></body></body></html>");
    }
    //------------------------------------------------------------------------------------------------------------

    //--------------------------------------------------------------------------------------------------------------

    /**
     * Сохраняет файл на сервере, в папке upload.
     * Сама папка должна быть уже создана.
     *
     * @param item
     * @throws Exception
     */
    private Boolean processUploadedFile(FileItem item, String sessionId) throws Exception {

        String Filename =item.getName();
        File uploadetFile = new File(home_pth+"//"+sessionId+"//"+ Filename);
        //выбираем файлу имя пока не найдём свободное

        if(!uploadetFile.exists()&&Filename!="") {
            int ExtIndex = Filename.indexOf('.');
            if(Filename.substring(ExtIndex).toLowerCase().equals(".pdf")) {

                File folder = new File(home_pth+"//" + sessionId);
                if (!folder.exists())
                    folder.mkdirs();

                //создаём файл
                uploadetFile.createNewFile();
                //записываем в него данные
                item.write(uploadetFile);
                return true;
            }

        }
        return false;
    }

    /**
     * Выводит на консоль имя параметра и значение
     * @param item
     */
    private void processFormField(FileItem item) {
        System.out.println(item.getFieldName()+"="+item.getString());
    }
    //-------------------------------------------------------
    void  nextPage(String mes1,String mes2, HttpServletRequest request,  HttpServletResponse response) throws IOException, ServletException {


        File folder = new File(home_pth+"//"+ sessionId);
        String mes3=Arrays.toString(folder.list());

        request.setAttribute("message1",mes1);
        request.setAttribute("message2", mes2);

        if(mes3.equals("null")) {

            request.setAttribute("message3","Пусто");
        }
        else
            request.setAttribute("message3",mes3);

        request.getRequestDispatcher("index.jsp").forward(request, response);
        //  response.getWriter().println(PageGenerator.instance().getPage("work.jsp",pageVariables));

    }


}
