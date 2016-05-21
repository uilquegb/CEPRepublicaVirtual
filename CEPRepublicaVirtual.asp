<%
Server.ScriptTimeout = 180
Class CEPRepublicaVirtual
    Dim fCEP, fUF, fCidade, fBairro, fTipoLogradouro, fLogradouro, fCEPUnico, fErro, fDescricaoErro
    
    Public Property Get CEP()
        CEP = fCEP
    End Property
    
    Public Property Let CEP(sCEP)
        fCEP = sCEP
    End Property
    
    Public Property Get UF()
         UF = fUF
    End Property
    
    Public Property Let UF(sUF)
         fUF = sUF
    End Property
    
    Public Property Get Cidade()
        Cidade = fCidade
    End Property
    
    Public Property Let Cidade(sCidade)
        fCidade = sCidade
    End Property
    
    Public Property Get Bairro()
        Bairro = fBairro
    End Property
    
    Public Property Let Bairro(sBairro)
        fBairro = sBairro
    End Property
    
    Public Property Get TipoLogradouro()
        TipoLogradouro = fTipoLogradouro
    End Property
    
    Public Property Let TipoLogradouro(sTipoLogradouro)
        fTipoLogradouro = sTipoLogradouro
    End Property
    
    Public Property Get Logradouro()
        Logradouro = fLogradouro
    End Property
    
    Public Property Let Logradouro(sLogradouro)
        fLogradouro = sLogradouro
    End Property
    
    Public Property Get CEPUnico()
        CEPUnico = fCEPUnico
    End Property
    
    Public Property Let CEPUnico(sCEPUnico)
        fCEPUnico = sCEPUnico
    End Property
    
    Public Property Get Erro()
        Erro = fErro
    End Property
    
    Public Property Get DescricaoErro()
        DescricaoErro = fDescricaoErro
    End Property
    
    Public Sub Buscar
        Dim resultado
        resultado = busca_cep(CEP)
        
        fErro = False
        fTipoLogradouro = ""
        fLogradouro = ""
        fBairro = ""
        fCidade = ""
        fUF = ""
        fCEPUnico = True
        fDescricaoErro = resultado( 3 )
        
        Select Case resultado( 2 )  
            Case "2"
                fCidade = resultado( 8 )
                fUF = resultado( 6 )
            Case "1"
                fTipoLogradouro = resultado( 12 )
                fLogradouro = resultado( 14 )
                fBairro = resultado( 10 )
                fCidade = resultado( 8 )
                fUF = resultado( 6 )
                fCEPUnico = False
            Case "0"
                fErro = True
        End Select  
    End Sub
    
    Private Function busca_cep(cep)   
    
        url = "http://republicavirtual.com.br/web_cep.php?cep="& cep &"&formato=query_string"  
        
        set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")    
        xmlhttp.open "GET", url, false    
        xmlhttp.send ""  
            
        xmlhttp_resultado = xmlhttp.responseText    
        set xmlhttp = nothing    
    
        arr_resultado = split(xmlhttp_resultado, "&")   
    
        Dim resultado(7)   
        
        For i = LBound(arr_resultado) To UBound(arr_resultado)   
            resultado(i) = arr_resultado(i) 
        Next   
    
        arr = split(join(resultado, "="), "=")
    
        Dim arr_2(14)   
        
        For i = LBound(arr) To UBound(arr)   
    
            arr_2(i) = replace(arr( i ), "+", " ")
            arr_2(i) = replace(arr_2(i), "%E1", "á")
            arr_2(i) = replace(arr_2(i), "%C1", "Á")
            arr_2(i) = replace(arr_2(i), "%E3", "ã")
            arr_2(i) = replace(arr_2(i), "%C3", "Ã")
            arr_2(i) = replace(arr_2(i), "%E2", "â")
            arr_2(i) = replace(arr_2(i), "%C2", "Â")
            arr_2(i) = replace(arr_2(i), "%E9", "é")
            arr_2(i) = replace(arr_2(i), "%C9", "É")
            arr_2(i) = replace(arr_2(i), "%E8", "è")
            arr_2(i) = replace(arr_2(i), "%C8", "È")
            arr_2(i) = replace(arr_2(i), "%EA", "ê")
            arr_2(i) = replace(arr_2(i), "%CA", "Ê")
            arr_2(i) = replace(arr_2(i), "%ED", "í")
            arr_2(i) = replace(arr_2(i), "%CD", "Í")
            arr_2(i) = replace(arr_2(i), "%EC", "ì")
            arr_2(i) = replace(arr_2(i), "%CC", "Ì")
            arr_2(i) = replace(arr_2(i), "%EE", "î")
            arr_2(i) = replace(arr_2(i), "%CE", "Î")
            arr_2(i) = replace(arr_2(i), "%F3", "ó")
            arr_2(i) = replace(arr_2(i), "%D3", "Ó")
            arr_2(i) = replace(arr_2(i), "%F2", "ò")
            arr_2(i) = replace(arr_2(i), "%D2", "Ò")
            arr_2(i) = replace(arr_2(i), "%F4", "ô")
            arr_2(i) = replace(arr_2(i), "%D4", "Ô")
            arr_2(i) = replace(arr_2(i), "%F5", "õ")
            arr_2(i) = replace(arr_2(i), "%D5", "Õ")
            arr_2(i) = replace(arr_2(i), "%FA", "ú")
            arr_2(i) = replace(arr_2(i), "%DA", "Ú")
            arr_2(i) = replace(arr_2(i), "%F9", "ù")
            arr_2(i) = replace(arr_2(i), "%D9", "Ù")
            arr_2(i) = replace(arr_2(i), "%FB", "û")
            arr_2(i) = replace(arr_2(i), "%DB", "Û")
            arr_2(i) = replace(arr_2(i), "%u0169", "ũ")
            arr_2(i) = replace(arr_2(i), "%u0168", "Ũ")
            arr_2(i) = replace(arr_2(i), "%E7", "ç")
            arr_2(i) = replace(arr_2(i), "%C7", "Ç")
            arr_2(i) = replace(arr_2(i), "/", "/")
            arr_2(i) = replace(arr_2(i), "/", "/")
            arr_2(i) = replace(arr_2(i), "%20", " ")
            arr_2(i) = replace(arr_2(i), "+", " ")
        Next       
        
        busca_cep = arr_2   
    End Function
End Class
%>
