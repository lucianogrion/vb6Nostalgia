Attribute VB_Name = "clsNegocio"
Option Explicit

'TODO:Agregar on error

Public Function RelevoArchivo(pathCompleto As String, fechaMod As Date) As Boolean
    Dim con As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rta As Boolean
    Dim sql As String
    
    rta = False
    Call AbrirConexionDB(con)
    Set rst = New ADODB.Recordset
    rst.Open "SELECT pathCompleto,intFecha FROM Archivo where pathCompleto='" + pathCompleto + "'", con, adOpenDynamic, adLockOptimistic
    
    If Not rst.EOF Then
        'Existe -> hay que validar la fecha
        If (rst.Fields("intFecha") <> Format(fechaMod, "MM/DD/YYYY HH:MM:SS")) Then
            sql = ""
            sql = sql + "update archivo set intFecha='" & Format(fechaMod, "MM/DD/YYYY HH:MM:SS") & "' , cambio=1 where pathCompleto='" & pathCompleto & "'"
            con.Execute (sql)
            rta = True
        End If
    Else
        sql = ""
        sql = sql & "Insert into Archivo(pathCompleto,intFecha,cambio) values ('" & pathCompleto & "','" & Format(fechaMod, "MM/DD/YYYY HH:MM:SS") & "',1)"
        con.Execute (sql)
        rta = True
    End If
    
    rst.Close
    Set rst = Nothing
    con.Close
    Set con = Nothing
    
    RelevoArchivo = rta
End Function


Public Sub MarcarArchivo(pathCompleto As String, fechaMod As Date)
    Dim con As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rta As Boolean
    Dim sql As String
    
    rta = False
    Call AbrirConexionDB(con)
    Set rst = New ADODB.Recordset
    rst.Open "SELECT pathCompleto,intFecha FROM Archivo where pathCompleto='" + pathCompleto + "'", con, adOpenDynamic, adLockOptimistic
    
    sql = ""
    sql = sql + "update archivo set intFecha='" & Format(fechaMod, "MM/DD/YYYY HH:MM:SS") & "' , cambio=0 where pathCompleto='" & pathCompleto & "'"
    con.Execute (sql)
    
    rst.Close
    Set rst = Nothing
    con.Close
    Set con = Nothing
    
End Sub



Public Sub CargarflxFacturacion(cuil As String, flxFacturacion As MSFlexGrid, aniol As Integer, mesl As Integer)

        Dim con As ADODB.Connection
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        
        Call LimpiarFlx(flxFacturacion)
        Call AbrirConexionDB(con)
        
        Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
        
        
        Dim sql As String
        sql = ""
        sql = sql + "Select "
        sql = sql + " FechaIng,nroFactura,mesl as MesFact,aniol as AnioFact,Cantidad"
        sql = sql + " FROM Facturacion"
        sql = sql + "  where CUIL='" & cuil & "'"
        If aniol <> 0 Then
            sql = sql & " AND aniol=" & aniol
            sql = sql & " AND mesl=" & mesl
        End If
        sql = sql + " ORDER BY mesl,aniol,FechaIng"
        
        rst.Open sql, con, adOpenDynamic, adLockOptimistic
        
        If Not rst.EOF Then
            i = 0
            For Each Campo In rst.Fields
               flxFacturacion.TextMatrix(0, i) = Campo.Name
               i = i + 1
            Next
            
            Do While Not rst.EOF
                dato = ""
                dato = dato & rst("FechaIng").Value & vbTab
                dato = dato & rst("nroFactura").Value & vbTab
                dato = dato & rst("MesFact").Value & vbTab
                dato = dato & rst("AnioFact").Value & vbTab
                dato = dato & rst("Cantidad").Value & vbTab
                
                flxFacturacion.AddItem (dato)
                rst.MoveNext
            Loop
            
        Else
            'lblNombre.Caption = "No se encontro Facturacion para este evaluador"
        End If
        
        'Redimensionado
        flxFacturacion.ColWidth(0) = 2500
        flxFacturacion.ColWidth(1) = 1500
        
        
        rst.Close
        Set rst = Nothing
        
        con.Close
        Set con = Nothing
End Sub


Public Function CargarflxEvaluacion(cuil As String, flxEvaluaciones As MSFlexGrid, aniol As Integer, mesl As Integer) As String
    Dim con As ADODB.Connection
    Dim retorno As String
    Call LimpiarFlx(flxEvaluaciones)
    Call AbrirConexionDB(con)
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
        
        
    Dim sql As String
    
    sql = ""
    sql = sql + " SELECT "
    sql = sql + " Evaluacion.pathCompleto, "
    sql = sql + " Evaluacion.EvaluacionFecha, "
    sql = sql + " Evaluacion.Nro, "
    sql = sql + " Evaluacion.PersonaNombre, "
    sql = sql + " Evaluacion.DocumentoTipo, "
    sql = sql + " Evaluacion.DocumentoNro, "
    sql = sql + " Evaluacion.CUIL, "
    sql = sql + " Evaluacion.Direccion, "
    sql = sql + " Evaluacion.Localidad, "
    sql = sql + " Evaluacion.Provincia, "
    sql = sql + " Evaluacion.ProvinciaCodigo, "
    sql = sql + " Evaluacion.Municipio, "
    sql = sql + " Evaluacion.EvaluadorNombe "
    sql = sql + " FROM Evaluacion "
    sql = sql + " WHERE EvaluadorCUIL='" & cuil & "' "
    If aniol <> 0 Then
        sql = sql & " AND year(Evaluacion.evaluacionfecha)=" & aniol
        sql = sql & " AND month(Evaluacion.evaluacionfecha)=" & mesl
    End If
        
    sql = sql + " ORDER BY EvaluacionFecha,nro"
    
    rst.Open sql, con, adOpenDynamic, adLockOptimistic
    
    If Not rst.EOF Then
        
        retorno = rst.Fields("EvaluadorNombe").Value
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        i = 0
        flxEvaluaciones.Row = 0
        For Each Campo In rst.Fields
            ' -- Agrega las columnas
           flxEvaluaciones.TextMatrix(0, i) = Campo.Name
           i = i + 1
        Next
            
        Do While Not rst.EOF

            dato = ""
            dato = dato & rst("pathCompleto").Value & vbTab
            dato = dato & rst("EvaluacionFecha").Value & vbTab
            dato = dato & rst("Nro").Value & vbTab
            dato = dato & rst("PersonaNombre").Value & vbTab
            dato = dato & rst("DocumentoTipo").Value & vbTab
            dato = dato & rst("DocumentoNro").Value & vbTab
            dato = dato & rst("CUIL").Value & vbTab
            dato = dato & rst("Direccion").Value & vbTab
            dato = dato & rst("Localidad").Value & vbTab
            dato = dato & rst("Provincia").Value & vbTab
            dato = dato & rst("ProvinciaCodigo").Value & vbTab
            dato = dato & rst("Municipio").Value & vbTab
            dato = dato & rst("EvaluadorNombe").Value & vbTab

            flxEvaluaciones.AddItem (dato)
            rst.MoveNext
        Loop
        
    Else
        'lblNombre.Caption = "No se encontraron evaluaciones para este evaluador"
    End If
    
    'Redimensionado
    flxEvaluaciones.ColWidth(0) = 2000
    flxEvaluaciones.ColAlignment(0) = vbAlignRight
    
    flxEvaluaciones.ColWidth(1) = 1500
    flxEvaluaciones.ColWidth(2) = 500
    flxEvaluaciones.ColWidth(3) = 2000
    flxEvaluaciones.ColWidth(4) = 500
    'flxEvaluaciones.ColWidth(5) = 500
    flxEvaluaciones.ColWidth(6) = 1500
    flxEvaluaciones.ColWidth(8) = 2000
    
    
        
    rst.Close
    Set rst = Nothing
    
    con.Close
    Set con = Nothing
        
    CargarflxEvaluacion = retorno
End Function


Public Sub CargarflxConciliacion(cuil As String, flxConciliacion As MSFlexGrid)
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        Dim con As ADODB.Connection
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        Dim rst As ADODB.Recordset
        Dim sql As String
        
        Call LimpiarFlx(flxConciliacion)
        Call AbrirConexionDB(con)
        
        sql = ""
        sql = sql + " SELECT  "
        sql = sql + " meses.aniol as Año ,  "
        sql = sql + " meses.mesl as Mesl, "
        sql = sql + " meses.Desc as Mes, "
        sql = sql + " Fact_AgrupMes.Total AS TOTALFACTURADO, "
        sql = sql + " Eval_AgrupMes.cantidad AS TOTALEVALUADO "
        sql = sql + " FROM (meses LEFT JOIN Fact_AgrupMes ON (meses.mesl = Fact_AgrupMes.mesl) AND (meses.aniol = Fact_AgrupMes.aniol))  "
        sql = sql + " LEFT JOIN Eval_AgrupMes ON (meses.mesl = Eval_AgrupMes.mes) AND (meses.aniol = Eval_AgrupMes.Anio) "
'        sql = sql & " Where Fact_AgrupMes.CUIL='" & Me.txtCUIL & "' "
'        sql = sql & " AND Eval_AgrupMes.evaluadorCUIL='" & Me.txtCUIL & "' "
        sql = sql + " ORDER BY meses.aniol,meses.mesl  ;  "
         
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        cmd.CommandText = sql
        
        cmd.Parameters.Append cmd.CreateParameter("micuil", adVarChar, adParamInput, 50, cuil)
         
        Set rst = New ADODB.Recordset
        Set rst = cmd.Execute  '.Open sql, con, adOpenDynamic, adLockOptimistic
        
        If Not rst.EOF Then
            i = 0
            For Each Campo In rst.Fields
               flxConciliacion.TextMatrix(0, i) = Campo.Name
               i = i + 1
            Next
            
            Do While Not rst.EOF
                dato = ""
                dato = dato & rst("Año").Value & vbTab
                dato = dato & rst("mesl").Value & vbTab
                dato = dato & rst("Mes").Value & vbTab
                dato = dato & rst("TOTALFACTURADO").Value & vbTab
                dato = dato & rst("TOTALEVALUADO").Value & vbTab
                
                flxConciliacion.AddItem (dato)
                rst.MoveNext
            Loop
            
        Else
            'lblNombre.Caption = "No se encontro Facturacion para este evaluador"
        End If
        
        'Redimensionado
        flxConciliacion.ColWidth(3) = 2800
        flxConciliacion.ColWidth(4) = 2750
        
        rst.Close
        Set rst = Nothing
        
        con.Close
        Set con = Nothing
        
End Sub



Public Sub CargarComboEvaluador(combo As ComboBox)
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        Dim con As ADODB.Connection
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        Dim rst As ADODB.Recordset
        Dim sql As String
        

        Call AbrirConexionDB(con)
        
        sql = ""
        sql = sql + " select distinct  "
        sql = sql + " evaluadornombe , evaluadorcuil"
        sql = sql + " From evaluacion"
        sql = sql + " order by evaluadorNombe "
        
        Set rst = New ADODB.Recordset
        rst.Open sql, con, adOpenDynamic, adLockOptimistic
        
        If Not rst.EOF Then
        
            Do While Not rst.EOF
                If (rst("evaluadornombe") <> "") Then
                    combo.AddItem rst("evaluadornombe").Value + "-" + rst("evaluadorCUIL").Value
                End If
                rst.MoveNext
            Loop
        Else
            '?
        End If
        
        
        
        rst.Close
        Set rst = Nothing
        
        con.Close
        Set con = Nothing
        
End Sub

Public Sub CargarComboTutor(combo As ComboBox)
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        Dim con As ADODB.Connection
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        Dim rst As ADODB.Recordset
        Dim sql As String
        

        Call AbrirConexionDB(con)
        
        sql = ""
        sql = sql + " select distinct  "
        sql = sql + " Tutor "
        sql = sql + " From evaluacion"
        sql = sql + " order by Tutor "
        
        Set rst = New ADODB.Recordset
        rst.Open sql, con, adOpenDynamic, adLockOptimistic
        
        If Not rst.EOF Then
        
            Do While Not rst.EOF
                If (rst("Tutor") <> "") Then
                    combo.AddItem rst("Tutor").Value + "-" + rst("Tutor").Value
                End If
                rst.MoveNext
            Loop
        Else
            '?
        End If
        
        
        
        rst.Close
        Set rst = Nothing
        
        con.Close
        Set con = Nothing
        
End Sub

Public Function CargarflxTutor(strTutor As String, flxEvaluaciones As MSFlexGrid, aniol As Integer, mesl As Integer) As String
    Dim con As ADODB.Connection
    Dim retorno As String
    Call LimpiarFlx(flxEvaluaciones)
    Call AbrirConexionDB(con)
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
        
        
    Dim sql As String
    
    sql = ""
    sql = sql + " SELECT "
    sql = sql + " Evaluacion.pathCompleto, "
    sql = sql + " Evaluacion.EvaluacionFecha, "
    sql = sql + " Evaluacion.Nro, "
    sql = sql + " Evaluacion.PersonaNombre, "
    sql = sql + " Evaluacion.DocumentoTipo, "
    sql = sql + " Evaluacion.DocumentoNro, "
    sql = sql + " Evaluacion.CUIL, "
    sql = sql + " Evaluacion.Direccion, "
    sql = sql + " Evaluacion.Localidad, "
    sql = sql + " Evaluacion.Provincia, "
    sql = sql + " Evaluacion.ProvinciaCodigo, "
    sql = sql + " Evaluacion.Municipio, "
    sql = sql + " Evaluacion.EvaluadorNombe "
    sql = sql + " FROM Evaluacion "
    sql = sql + " WHERE tutor='" & strTutor & "' "
    If aniol <> 0 Then
        sql = sql & " AND year(Evaluacion.evaluacionfecha)=" & aniol
        sql = sql & " AND month(Evaluacion.evaluacionfecha)=" & mesl
    End If
        
    sql = sql + " ORDER BY EvaluacionFecha,nro"
    
    rst.Open sql, con, adOpenDynamic, adLockOptimistic
    
    If Not rst.EOF Then
        
        retorno = rst.Fields("EvaluadorNombe").Value
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        i = 0
        flxEvaluaciones.Row = 0
        For Each Campo In rst.Fields
            ' -- Agrega las columnas
           flxEvaluaciones.TextMatrix(0, i) = Campo.Name
           i = i + 1
        Next
            
        Do While Not rst.EOF

            dato = ""
            dato = dato & rst("pathCompleto").Value & vbTab
            dato = dato & rst("EvaluacionFecha").Value & vbTab
            dato = dato & rst("Nro").Value & vbTab
            dato = dato & rst("PersonaNombre").Value & vbTab
            dato = dato & rst("DocumentoTipo").Value & vbTab
            dato = dato & rst("DocumentoNro").Value & vbTab
            dato = dato & rst("CUIL").Value & vbTab
            dato = dato & rst("Direccion").Value & vbTab
            dato = dato & rst("Localidad").Value & vbTab
            dato = dato & rst("Provincia").Value & vbTab
            dato = dato & rst("ProvinciaCodigo").Value & vbTab
            dato = dato & rst("Municipio").Value & vbTab
            dato = dato & rst("EvaluadorNombe").Value & vbTab

            flxEvaluaciones.AddItem (dato)
            rst.MoveNext
        Loop
        
    Else
        'lblNombre.Caption = "No se encontraron evaluaciones para este evaluador"
    End If
    
    'TODO: Falta agrupar por diferentes CUILES por tutor
    
    
    'Redimensionado
    flxEvaluaciones.ColWidth(0) = 2000
    flxEvaluaciones.ColAlignment(0) = vbAlignRight
    
    flxEvaluaciones.ColWidth(1) = 1500
    flxEvaluaciones.ColWidth(2) = 500
    flxEvaluaciones.ColWidth(3) = 2000
    flxEvaluaciones.ColWidth(4) = 500
    'flxEvaluaciones.ColWidth(5) = 500
    flxEvaluaciones.ColWidth(6) = 1500
    flxEvaluaciones.ColWidth(8) = 2000
    
    
        
    rst.Close
    Set rst = Nothing
    
    con.Close
    Set con = Nothing
        
    CargarflxTutor = retorno
End Function


Public Sub CargarflxFacturacionTutor(strTutor As String, flxFacturacion As MSFlexGrid, aniol As Integer, mesl As Integer)

        Dim con As ADODB.Connection
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        
        Call LimpiarFlx(flxFacturacion)
        Call AbrirConexionDB(con)
        
        Dim rst As ADODB.Recordset
        Set rst = New ADODB.Recordset
        
        
        Dim sql As String
        sql = ""
        sql = sql + "Select "
        sql = sql + " FechaIng,nroFactura,mesl as MesFact,aniol as AnioFact,Cantidad"
        sql = sql + " FROM FacturacionTutor"
        sql = sql + "  where Tutor='" & strTutor & "'"
        If aniol <> 0 Then
            sql = sql & " AND aniol=" & aniol
            sql = sql & " AND mesl=" & mesl
        End If
        sql = sql + " ORDER BY mesl,aniol,FechaIng"
        
        rst.Open sql, con, adOpenDynamic, adLockOptimistic
        
        If Not rst.EOF Then
            i = 0
            For Each Campo In rst.Fields
               flxFacturacion.TextMatrix(0, i) = Campo.Name
               i = i + 1
            Next
            
            Do While Not rst.EOF
                dato = ""
                dato = dato & rst("FechaIng").Value & vbTab
                dato = dato & rst("nroFactura").Value & vbTab
                dato = dato & rst("MesFact").Value & vbTab
                dato = dato & rst("AnioFact").Value & vbTab
                dato = dato & rst("Cantidad").Value & vbTab
                
                flxFacturacion.AddItem (dato)
                rst.MoveNext
            Loop
            
        Else
            'lblNombre.Caption = "No se encontro Facturacion para este evaluador"
        End If
        
        'Redimensionado
        flxFacturacion.ColWidth(0) = 2500
        flxFacturacion.ColWidth(1) = 1500
        
        
        rst.Close
        Set rst = Nothing
        
        con.Close
        Set con = Nothing
End Sub

Public Sub CargarflxConciliacionTutor(tutor As String, flxConciliacion As MSFlexGrid)
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        Dim con As ADODB.Connection
        Dim dato As String
        Dim Campo As ADODB.Field
        Dim i As Integer
        Dim rst As ADODB.Recordset
        Dim sql As String
        
        Call LimpiarFlx(flxConciliacion)
        Call AbrirConexionDB(con)
        
        sql = ""
        sql = sql + " SELECT  "
        sql = sql + " meses.aniol as Año ,  "
        sql = sql + " meses.mesl as Mesl, "
        sql = sql + " meses.Desc as Mes, "
        sql = sql + " Fact_AgrupMes.Total AS TOTALFACTURADO, "
        sql = sql + " Eval_AgrupMes.cantidad AS TOTALEVALUADO "
        sql = sql + " FROM (meses LEFT JOIN Fact_AgrupMesTutor ON (meses.mesl = Fact_AgrupMes.mesl) AND (meses.aniol = Fact_AgrupMes.aniol))  "
        sql = sql + " LEFT JOIN Tutorias_AgrupMes ON (meses.mesl = Eval_AgrupMes.mes) AND (meses.aniol = Eval_AgrupMes.Anio) "
'        sql = sql & " Where Fact_AgrupMes.CUIL='" & Me.txtCUIL & "' "
'        sql = sql & " AND Eval_AgrupMes.evaluadorCUIL='" & Me.txtCUIL & "' "
        sql = sql + " ORDER BY meses.aniol,meses.mesl  ;  "
         
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        cmd.CommandText = sql
        
        cmd.Parameters.Append cmd.CreateParameter("micuil", adVarChar, adParamInput, 50, tutor)
         
        Set rst = New ADODB.Recordset
        Set rst = cmd.Execute  '.Open sql, con, adOpenDynamic, adLockOptimistic
        
        If Not rst.EOF Then
            i = 0
            For Each Campo In rst.Fields
               flxConciliacion.TextMatrix(0, i) = Campo.Name
               i = i + 1
            Next
            
            Do While Not rst.EOF
                dato = ""
                dato = dato & rst("Año").Value & vbTab
                dato = dato & rst("mesl").Value & vbTab
                dato = dato & rst("Mes").Value & vbTab
                dato = dato & rst("TOTALFACTURADO").Value & vbTab
                dato = dato & rst("TOTALEVALUADO").Value & vbTab
                
                flxConciliacion.AddItem (dato)
                rst.MoveNext
            Loop
            
        Else
            'lblNombre.Caption = "No se encontro Facturacion para este evaluador"
        End If
        
        'Redimensionado
        flxConciliacion.ColWidth(3) = 2800
        flxConciliacion.ColWidth(4) = 2750
        
        rst.Close
        Set rst = Nothing
        
        con.Close
        Set con = Nothing
        
End Sub

