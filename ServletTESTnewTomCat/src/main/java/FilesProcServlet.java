import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;
import com.itextpdf.text.pdf.parser.SimpleTextExtractionStrategy;
import com.itextpdf.text.pdf.parser.TextExtractionStrategy;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;

import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigDecimal;
import java.sql.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static org.apache.poi.ss.formula.BaseFormulaEvaluator.evaluateAllFormulaCells;
import static org.apache.poi.ss.usermodel.BorderStyle.MEDIUM;
import static org.apache.poi.ss.usermodel.BorderStyle.THICK;
import static org.apache.poi.ss.usermodel.BorderStyle.THIN;

/**
 * Created by Damned on 14.08.2017.
 */
public class FilesProcServlet extends HttpServlet {
    Connection connect=null;
    String sessionId="";
    String home_pth=System.getProperty("user.dir")+"/out";

    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException
    {
        request.setAttribute("message1","Начало работы");
        request.setAttribute("message2","");
        request.setAttribute("message3","");
        doPost(request,response);
    }
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException
    {

        // GПОлчение идентификатора сессии
        sessionId = request.getSession().getId();
        ArrayList<String> invArr=new ArrayList<String>();
        // получение списка загруженных файлов
        File folder = new File(home_pth+"//"+ sessionId);
        String Flist[] = folder.list();

        if(Flist==null)
        {
            if(((String)request.getAttribute("message3")).compareTo("")==0)
                request.setAttribute("message3","Нет загруженых файлов для обработки");
            request.getRequestDispatcher("index.jsp").forward(request, response);
            return;
        }
        //обрабатываем файл из списка файлов
        for(String s:Flist) {
            // проверяем действительно ли PDF
            int ExtIndex = s.indexOf('.');
            if(s.substring(ExtIndex).toLowerCase().equals(".pdf")) {
                System.out.println("File to process:" + s);
                // Обрабатываем файл и получаем номер счета
                String invStr = ProccessPDFfile(folder.toString() + "//" + s);
                // запоминаем какие счета мы обработали
                if(!invStr.equals(""))invArr.add(invStr);
            }
        }
        // удвляем папку для временного хранения pdf
        folder.delete();

        System.out.println("Processed invoices:"+ String.join(", ", invArr));
        String T1_statement="";
        String T2_statement="";
// формирование запроса на полчения выборки отчета по пакету в баху данных
        if(invArr.size()>0) {
            T1_statement = makeStatementT1(invArr);
            T2_statement=makeStatementT2(invArr);
        }
        makeEXEL(T1_statement,T2_statement,response);

    }
    //-------------------------------------------------------------------------------------------------------
// Метод выделяем нужную информацию из PDF файла и щаписывает ее в базу данных
    private String ProccessPDFfile(String fName) throws IOException {

        Statement statement = null;
        PdfReader reader = new PdfReader(fName);
        DecimalFormat pos_num_formt =new DecimalFormat("000");
        int pos_num=0;
        String pos_num_str="";
        String invoice="";
        String t_weight="";
        String n_weight="";
        String n_pack="";
        String textAll="";
        int ind=0;
        int ind2=0;
        int ind3=0;
        // подключаемся к БД
        try {
            connect=Dbconnect();
            statement =connect.createStatement();
        } catch (SQLException e) {
            e.printStackTrace();
        //    endConnection();
            return null;
        }

        if(connect!=null)
        {
            //      System.out.println("Proccessing PDF file:"+fName);
// Открываем PDF файл и читаем его постранично

            for (int i = 1; i <= reader.getNumberOfPages(); i++) {
                TextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                String text = PdfTextExtractor.getTextFromPage(reader, i, strategy);
// ищем поля для заполнения таблицы invoices на первой страницы PDF
                if(i==1)
                {
                    //Ищем номер счета
                    pos_num= text.indexOf("Customs Invoice-no:",0);
                    invoice=text.substring(pos_num+20,text.indexOf("\n",pos_num+20));
                    //  System.out.println("invoice:"+invoice);
                    //Ищем общую массу
                    pos_num= text.indexOf("Weight gross",0);
                    t_weight=text.substring(pos_num+12,text.indexOf(",",pos_num+12)+4);
                    // форматируем строку для преобразования в Double
                    t_weight=t_weight.replaceAll(" ","");
                    t_weight=t_weight.replace(".","");
                    t_weight=t_weight.replace(",",".");
                    //    System.out.println("Total weight:"+t_weight);
                    //Ищем  массу нетто
                    pos_num= text.indexOf("Weight net",0);
                    n_weight=text.substring(pos_num+10,text.indexOf(",",pos_num+10)+4);
                    // форматируем строку для преобразования в Double
                    n_weight=n_weight.replaceAll(" ","");
                    n_weight=n_weight.replace(".","");
                    n_weight=n_weight.replace(",",".");
                    //   System.out.println("Net weight:"+n_weight);
                    //Ищем  число упаковок
                    pos_num= text.indexOf("Packing",0);
                    n_pack=text.substring(pos_num+7,text.indexOf("Total No",pos_num));
                    n_pack=n_pack.replaceAll(" ","");
                    //    System.out.println("Packing:"+n_pack);

                    double temp1, temp2;
                    temp1=Double.parseDouble(t_weight);
                    temp2= Double.parseDouble(n_weight);
                    try {
                        statement.executeUpdate(
                                "insert into invoices values                "
                                        +"("+ invoice+","+t_weight+","+n_weight+","+n_pack+")");
                    } catch (SQLException e) {
                        e.printStackTrace();
                     //   endConnection();
                        return null;
                    }
                }
                else
                {
                    // объединяем остальные страницы в один кусок текста и удаляем колонтитулы
                    ind=0;
                    // Удаление колонтитулов
                    ind=text.indexOf("\n_________________________________________________________________________",ind);

                    if(ind>0)ind+=74;
                    ind2=0;
                    ind3=0;
                    ind2=text.indexOf("Interface GmbH",ind);
                    ind3=text.indexOf("ller Interface",ind);
                    // Склеиваем страницы
                    if(ind3<ind2&&ind3>0)
                        textAll+=text.substring(ind,ind3-7);
                    else
                        textAll+=text.substring(ind,ind2);
                }
            }
            //------------------------------------------------------------------------------------------
            // разбиваем текст на позиции и выделяем нужную информацию
            //разбиваем текст на позиции
            ind=0;
            pos_num=0;

            pos_num_str=pos_num_formt.format(pos_num+1)+"/";
            while((ind=textAll.indexOf(pos_num_str,ind))>0)
            {

                String code="";
                String name="";
                String quan="";
                String ed="";
                String w_brutto="";

                ind2=textAll.indexOf("weight of position",ind);
                ind2=textAll.indexOf("\n",ind2);
                String subtext=textAll.substring(ind,ind2);
                // выделяем информацию из позиции
                //   System.out.println("\n"+pos_num_str);
                // Ищем количество
                int ind_sub=0;
                int ind_sub2=0;
                ind_sub=subtext.indexOf("Item",ind_sub);
                ind_sub2=subtext.indexOf("\n",ind_sub)+1;
                ind_sub=subtext.indexOf(",",ind_sub2);
                quan=subtext.substring(ind_sub2,ind_sub+3);
                quan=quan.replaceAll(" ", "");
                //    System.out.println("Quantity:"+quan);
                // форматируем строку для преобразования в int
                quan=quan.replace(".", "");
                quan=quan.substring(0,quan.indexOf(","));
                //    System.out.println("QuantityF:"+quan);

                // ищим единицы измерения
                ind_sub2=subtext.indexOf(" ",ind_sub+4);
                ed=subtext.substring(ind_sub+4,ind_sub2);
                // System.out.println("Ed:"+ed);
                // ищим наименование
                ind_sub=subtext.indexOf("\n",ind_sub2);
                ind_sub2=subtext.indexOf("\n",ind_sub+1);
                name=subtext.substring(ind_sub+1,ind_sub2);
                //      System.out.println("Name:"+name);
                // Ищим код
                ind_sub=subtext.indexOf("customs tariff no.",ind_sub2);
                ind_sub2=subtext.indexOf("\n",ind_sub);
                code=subtext.substring(ind_sub+18,ind_sub2);
                code=code.replaceAll(" ", "");
                //     System.out.println("Code:"+code);
                // Ищем массу брутто
                ind_sub=subtext.indexOf("weight of position",ind_sub2);
                ind_sub2=subtext.indexOf("G",ind_sub);
                w_brutto=subtext.substring(ind_sub+18,ind_sub2);
                w_brutto=w_brutto.replaceAll(" ", "");
                //     System.out.println("Brutto:"+w_brutto);

                //  System.out.println(subtext);
                pos_num++;
                pos_num_str=pos_num_formt.format(pos_num+1)+"/";
                //    ind2=ind;

                // Запись полученых значений в базу данных

                try {
                    statement.executeUpdate(
                            "insert into goods(code,name,quan,ed,w_brutto,invoice) values                "
                                    +"('"+ code+"','"+name+"',"+quan+",'"+ed+"',"+w_brutto+",'"+invoice+"')");
                } catch (SQLException e) {
                    e.printStackTrace();
           //         endConnection();
                    return null;
                }

            }
        }

        reader.close();
        // удаляеми файл
        File f= new File(fName);
        f.delete();
        // Возвращаем номер счета
        return invoice;

    }
    //-------------------------------------------------------------------------------------------------------
    // метод устанавливает соединение с базой данных
    Connection Dbconnect()
    {
        Connection connection=connect;
        if(connect==null) {
            String userName = "root";
            String password = "admin";
            String url = "jdbc:mysql://localhost:3306/test?useSSL=false";

            try {
                Class.forName("com.mysql.jdbc.Driver").newInstance();
                connection = DriverManager.getConnection(url, userName, password);
            } catch (Exception e) {
                System.out.println("Something wrong with driver");
                return null;
            }
        }
        return connection;
    }
    //-------------------------------------------------------------------------------------------------------
    // метод конструирует строку запроса в БД для получения итоговой таблицы 1
    String makeStatementT1(ArrayList<String> invL)
    {
        String st="";
        String invs=" or invoice=";
        st+="SELECT `code` as cod, `name` as nam, sum(quan)'sum', sum(w_brutto)'w_brutto', ed,"+
                "(SELECT GROUP_CONCAT( DISTINCT invoice) from goods where `code`=cod and `name`=nam ) 'invoice'"
                +" FROM goods where invoice="+ invL.get(0);;
        for(int i=1;i<invL.size();i++)
        {
            st+= invs+ invL.get(i);
        }
        st+=" GROUP BY `code`,`name`  ORDER BY `name`,invoice,`code`";

//        System.out.println(st);;
        return st;
    }
    //-------------------------------------------------------------------------------------------------------
    // метод конструирует строку запроса в БД для получения итоговой таблицы 1
    String makeStatementT2(ArrayList<String> invL)
    {
        String st="";
        String invs=" or invoice=";
        st+="SELECT * FROM invoices where invoice="+ invL.get(0);;
        for(int i=1;i<invL.size();i++)
        {
            st+= invs+ invL.get(i);
        }
        return st;
    }
    //-------------------------------------------------------------------------------------------------------
    // Метод создает 2 запроса в БД и на их основе создает EXCEL файл с результатами
    void makeEXEL(String T1_statement, String T2_statement, HttpServletResponse response)
    {

        Statement statement = null;
        ResultSet resultset=null;
        int row_num=1;

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();

        HSSFRow row;
        HSSFCell cell;
        // создание стиля ячейки
        HSSFCellStyle style1 = wb.createCellStyle();
        HSSFCellStyle style2 = wb.createCellStyle();
        BorderStyle st1=THIN;
        style1.setBorderBottom(st1);
        style1.setBorderRight(st1);
        style1.setBorderTop(st1);
        style1.setBorderLeft(st1);
        //-------------------
        BorderStyle st2=MEDIUM;
        style2.setBorderBottom(st2);
        style2.setBorderRight(st2);
        style2.setBorderTop(st2);
        style2.setBorderLeft(st2);
        //Запрос в БД для полчуения данных для Таблицы 1
        try {
            if(connect==null)
                connect=Dbconnect();

            statement =connect.createStatement();
            resultset = statement.executeQuery(T1_statement);
        }    catch (SQLException e) {
            e.printStackTrace();
        //    endConnection();
            return ;
        }

        // формирование шапки первой таблицы
        row = sheet.createRow(0);
        cell=row.createCell(0);
        cell.setCellStyle(style2);
        cell.setCellValue("Код");
        cell=row.createCell(1);
        cell.setCellStyle(style2);
        cell.setCellValue("Наименование товара");
        cell=row.createCell(2);
        cell.setCellStyle(style2);
        cell.setCellValue("Кол-во");
        cell=row.createCell(3);
        cell.setCellStyle(style2);
        cell.setCellValue("Ед. изм");
        cell=row.createCell(4);
        cell.setCellStyle(style2);
        cell.setCellValue("Вес брутто, кг");
        cell=row.createCell(5);
        cell.setCellStyle(style2);
        cell.setCellValue("Счет №");
        try {
            while(resultset.next())
            {
                // заполнение таблицы 1
                row = sheet.createRow(row_num);
                cell=row.createCell(0);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getString("cod"));
                cell=row.createCell(1);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getString("nam"));
                cell=row.createCell(2);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getString("sum"));
                cell=row.createCell(3);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getString("ed"));
                cell=row.createCell(4);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getDouble ("w_brutto"));
                cell=row.createCell(5);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getString("invoice"));
                row_num++;
            }

        } catch (SQLException e) {
            e.printStackTrace();
        //    endConnection();
            return ;
        }
        row_num++;
//--------------------
        //Запрос в БД для полчуения данных для Таблицы 2
        try {
            if(connect==null)
                connect=Dbconnect();

            statement =connect.createStatement();
            resultset = statement.executeQuery(T2_statement);
        }    catch (SQLException e) {
            e.printStackTrace();
        }
        // формирование шапки первой таблицы
        row = sheet.createRow(row_num);
        cell=row.createCell(0);
        cell.setCellStyle(style2);
        cell.setCellValue("Счет №");
        cell=row.createCell(1);
        cell.setCellStyle(style2);
        cell.setCellValue("Packages");
        cell=row.createCell(2);
        cell.setCellStyle(style2);
        cell.setCellValue("Net weight");
        cell=row.createCell(3);
        cell.setCellStyle(style2);
        cell.setCellValue("Total weight");
        row_num++;
        int r_beg=row_num+1;
        try {

            while(resultset.next())
            {
                // заполнение таблицы 1
                row = sheet.createRow(row_num);
                cell=row.createCell(0);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getString("invoice"));
                cell=row.createCell(1);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getDouble("t_weight"));
                cell=row.createCell(2);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getDouble("n_weight"));
                cell=row.createCell(3);
                cell.setCellStyle(style1);
                cell.setCellValue(resultset.getInt("n_pack"));

                row_num++;
            }
            for( int colNum = 0; colNum<6;colNum++)
                wb.getSheetAt(0).autoSizeColumn(colNum);
        } catch (SQLException e) {
            e.printStackTrace();
       //     endConnection();
            return ;
        }
     //   endConnection();
        // запись формул для подсчета в файл
        row = sheet.createRow(row_num);
        cell=row.createCell(1);
        cell.setCellFormula ("sum(B"+r_beg+":B"+row_num+")");
        cell=row.createCell(2);
        cell.setCellFormula ("sum(C"+r_beg+":C"+row_num+")");
        cell=row.createCell(3);
        cell.setCellFormula ("sum(D"+r_beg+":D"+row_num+")");
        evaluateAllFormulaCells( wb);
//---------------------------------------------------------------------
        //выгрузка excel файла клиенту
        if(wb!=null) {
            //    FileInputStream inStream = new FileInputStream(wb);

            // obtains ServletContext
            ServletContext context = getServletContext();

            // gets MIME type of the file
            String mimeType = context.getMimeType("application/vnd.ms-excel ");
            if (mimeType == null) {
                // set to binary type if MIME mapping not found
                mimeType = "application/octet-stream";
            }
            System.out.println("MIME type: " + mimeType);

            // modifies response
            try {
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                wb.write(baos);

                response.setContentType(mimeType);
                response.setContentLength(baos.size());

                // forces download
                String headerKey = "Content-Disposition";
                String headerValue = String.format("attachment; filename=\"%s\"", "res.xls");
                response.setHeader(headerKey, headerValue);

                // obtains response's output stream
                OutputStream outStream = null;
                outStream = response.getOutputStream();
                outStream.write(baos.toByteArray());
                outStream.close();
                baos.close();

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    //------------------------------------------------------------------------
    // Метод разрывает соединение с базой и удаляет загруженные пользователем файлы
    void endConnection()
    {
        if(connect!=null)
        {
            try {
                connect.close();
            } catch (SQLException e) {
                e.printStackTrace();
                System.out.println("Cant close connection");
            }
            connect=null;
        }
    }
}

