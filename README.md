# CEPRepublicaVirtual

Class em ASP Clássico para buscar CEP utilizando o site [Republica virtual](http://republicavirtual.com.br)

O principal objetivo desse repositório é fornecer uma classe utilitária para quem quer calcular CEP no seu site feito em `Classic ASP` (ASP Clássico ou ASP 3.0)

    <%
    Dim oCep
    Set oCep = New CEPRepublicaVirtual
    oCep.CEP = "01001-001" 'Poderia ser 01001001
    oCep.Buscar

    Response.Write "oCep.UF " & oCep.UF & "<br />"
    Response.Write "oCep.Cidade " & oCep.Cidade & "<br />"
    Response.Write "oCep.Bairro " & oCep.Bairro & "<br />"
    Response.Write "oCep.Logradouro " & oCep.Logradouro & "<br />"
    Response.Write "oCep.TipoLogradouro " & oCep.TipoLogradouro & "<br />"
    Response.Write "oCep.CEPUnico " & oCep.CEPUnico & "<br />"
    Response.Write "oCep.CEP " & oCep.CEP & "<br />"
    Response.Write "oCep.DescricaoErro " & oCep.DescricaoErro & "<br />"
    Response.Write "oCep.Erro " & oCep.Erro & "<br />"
    %>