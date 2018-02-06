<html>
<head>
  <meta charset="utf-8">
  <title>File Upload </title>
</head>
<body>
<div style="text-align: center;">

               <% String outParameter = (String)request.getAttribute("message1");
                         if (outParameter != null){%>
                         <%=  outParameter %> <br>
                         <% } else { %> Начало работы<br> <% }%>


</div>
<form method="post" action="/upload" enctype="multipart/form-data">
  Выберите .pdf файл:<br>
  <input type="file" name="uploadFile" />
  <br/><br/>
  <input type="submit" value="Загрузить" />

</form>
<form method="post" action="/process">
  <input type="submit" value="Обработать" />

</form>
<div style="text-align: center;">


                <%  outParameter = (String)request.getAttribute("message2");
                           if (outParameter != null){%>
                           <%=  outParameter %> <br>
                           <% } else { %> <br> <% }%>
                <%  outParameter = (String)request.getAttribute("message3");
                           if (outParameter != null){%>
                           <%=  outParameter %> <br>
                           <% } else { %> <br> <% }%>
<button onclick="/index.jsp" onclick="window.open('/process');">Обработать</button><br>
<a href="/index.jsp" onclick="window.open('/process');">Обработать</a<br>

</div>
</body>
</html>