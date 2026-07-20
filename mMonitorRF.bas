Attribute VB_Name = "mMonitorF0"
Option Explicit

'============================================================
' MONITOR RF - FONDO CONSERVADOR - Modulo unico (v3)
' Fuentes + motor + metricas + historico + diario/backfill
'
' MACROS PARA EL USUARIO (asignar a botones):
'   CrearEsqueleto      -> crea/repara todas las hojas
'   Inventariar         -> lista observaciones posibles
'   CorrerDiario        -> calcula la ultima observacion (boton 1)
'   CorrerBackfill      -> reconstruye historico (boton 2)
'
' REGLA DE PAREO (cerrada):
'   Observacion fechada por el VECTOR. Cartera = FMS anterior
'   (desfase Config B9: 1 = t-2/t-1, 0 = misma fecha).
'   Misma regla en diario y backfill.
'
' UNIVERSOS (cerrado):
'   A Negociable (sin depositos): spread, rating, concentracion,
'     YTM negociable, contribucion por instrumento
'   B Total (con depositos): YTM fondo, duracion, vencimientos,
'     liquidez. Depositos: tasa de hoja Depositos, duracion
'     interpolada de CP Duracion.
'============================================================

Private Const SH_CFG   As String = "Config"
Private Const SH_DEP   As String = "Depositos"
Private Const SH_CUR   As String = "CP Duracion"
Private Const SH_INV   As String = "Inventario"
Private Const SH_CART  As String = "Cartera F0"
Private Const SH_YLD   As String = "Yield"
Private Const SH_CAL   As String = "Calidad"
Private Const SH_CON   As String = "Concentracion"
Private Const SH_VEN   As String = "Vencimientos"
Private Const SH_CTR   As String = "Contribucion"
Private Const SH_HIS   As String = "Historico"

' Columnas de la matriz de cartera en memoria
Private Const C_COD = 1      ' codigo (cruce)
Private Const C_EMI = 2      ' emisor
Private Const C_AC = 3       ' asset class
Private Const C_CAT = 4      ' categoria (3 grupos)
Private Const C_MON = 5      ' moneda
Private Const C_CANT = 6     ' cantidad
Private Const C_MTO = 7      ' monto MM
Private Const C_RAT = 8      ' rating
Private Const C_VCTO = 9     ' vencimiento (fecha o vacio)
Private Const C_DIAS = 10    ' dias a vcto
Private Const C_YTW = 11     ' ytw (%)
Private Const C_SPR = 12     ' spread pbs
Private Const C_DUR = 13     ' duracion
Private Const C_VEC = 14     ' 1 si esta en vector
Private Const C_DEP = 15     ' 1 si es deposito
Private Const C_NCOLS = 15

'============================================================
' CONFIG Y UTILIDADES
'============================================================
Private Function Cfg(ByVal celda As String) As String
    Cfg = Trim$(CStr(ThisWorkbook.Worksheets(SH_CFG).Range(celda).Value))
End Function

Private Function CfgFecha(ByVal celda As String) As Date
    Dim v As Variant
    v = ThisWorkbook.Worksheets(SH_CFG).Range(celda).Value
    If IsDate(v) Then CfgFecha = CDate(v) Else CfgFecha = 0
End Function

Private Function CfgNum(ByVal celda As String, ByVal porDefecto As Double) As Double
    Dim v As Variant
    v = ThisWorkbook.Worksheets(SH_CFG).Range(celda).Value
    If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then CfgNum = CDbl(v) Else CfgNum = porDefecto
End Function

Private Function ConBarra(ByVal ruta As String) As String
    If Len(ruta) = 0 Then ConBarra = "": Exit Function
    If Right$(ruta, 1) <> "\" Then ruta = ruta & "\"
    ConBarra = ruta
End Function

Private Function Hoja(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    End If
    Set Hoja = ws
End Function

Private Function UltFila(ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then UltFila = 0 Else UltFila = c.Row
End Function

'============================================================
' ESQUELETO
'============================================================
Public Sub CrearEsqueleto()
    Dim ws As Worksheet

    Set ws = Hoja(SH_CFG)
    With ws
        .Range("A1").Value = "CARPETA FMS"
        .Range("A2").Value = "CARPETA VECTOR"
        .Range("A3").Value = "CARPETA HISTORICO (copias)"
        .Range("A4").Value = "FECHA INICIO backfill (vacio = todo)"
        .Range("A5").Value = "FECHA FIN backfill (vacio = ultimo)"
        .Range("A6").Value = "BASE DIAS"
        .Range("A7").Value = "CODIGO FONDO"
        .Range("A8").Value = "COL CRUCE FMS (7=G nemonico)"
        .Range("A9").Value = "DESFASE FMS (1=t-2/t-1, 0=misma fecha)"
        If Len(Cfg("B6")) = 0 Then .Range("B6").Value = 365
        If Len(Cfg("B7")) = 0 Then .Range("B7").Value = 9
        If Len(Cfg("B8")) = 0 Then .Range("B8").Value = 7
        If Len(Cfg("B9")) = 0 Then .Range("B9").Value = 1
        .Range("A1:A9").Font.Bold = True
        .Range("A:A").ColumnWidth = 36
        .Range("B:B").ColumnWidth = 45
        .Cells.Font.Name = "Arial": .Cells.Font.Size = 8
        .Range("B4:B5").NumberFormat = "dd/mm/yyyy"
    End With

    Set ws = Hoja(SH_DEP)
    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then
        ws.Range("A1").Value = "Pegar aqui la hoja Depositos de Duracion.xlsm  (C=codigo SBS, F=tasa pactada)"
        ws.Range("A1").Font.Italic = True
    End If

    Set ws = Hoja(SH_CUR)
    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then
        ws.Range("A1").Value = "DIAS": ws.Range("B1").Value = "DURACION"
        ws.Range("A2").Value = "Pegar aqui la curva de CP Duracion de Duracion.xlsm (A=dias, B=duracion, ascendente)"
        ws.Range("A2").Font.Italic = True
    End If

    Set ws = Hoja(SH_HIS)
    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then EncabezadoHistorico ws

    Hoja(SH_INV): Hoja(SH_CART): Hoja(SH_YLD): Hoja(SH_CAL)
    Hoja(SH_CON): Hoja(SH_VEN): Hoja(SH_CTR)

    MsgBox "Esqueleto listo." & vbCrLf & _
           "1) Llena rutas en Config (B1, B2, B3)." & vbCrLf & _
           "2) Pega Depositos y CP Duracion desde Duracion.xlsm." & vbCrLf & _
           "3) Corre Inventariar.", vbInformation
End Sub

Private Sub EncabezadoHistorico(ws As Worksheet)
    ws.Range("A1:AA1").Value = Array("FECHA", "DIAS", "VALCART MM", _
        "YTM FONDO", "YTM NEGOC", "YTM BONOS", "YTM PC", "YTM DEP", _
        "SPR TOT", "SPR BONOS", "SPR PC", _
        "DUR TOT", "DUR BONOS", "DUR PC", "DUR DEP", _
        "PCT IG", "RATING PROM", "HHI", "TOP5", "TOP10", _
        "VENC 30D", "VENC 90D", "CARRY DIA", "CARRY MTD", "CARRY YTD", _
        "N INSTR", "N SIN MATCH")
    FormatoInstitucional ws.Range("A1:AA1")
End Sub

'============================================================
' PARSER DE NOMBRES Y MAPAS DE ARCHIVOS
'============================================================
Private Function Fecha8(ByVal nombre As String) As Date
    Dim i As Long, s As String, a As Integer, m As Integer, d As Integer
    Fecha8 = 0
    For i = 1 To Len(nombre) - 7
        s = Mid$(nombre, i, 8)
        If EsSoloDigitos(s) Then
            a = CInt(Left$(s, 4)): m = CInt(Mid$(s, 5, 2)): d = CInt(Right$(s, 2))
            If a >= 2000 And a <= 2100 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                On Error Resume Next
                Fecha8 = DateSerial(a, m, d)
                On Error GoTo 0
                If Fecha8 > 0 Then Exit Function
            End If
        End If
    Next i
End Function

Private Function EsSoloDigitos(ByVal s As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then EsSoloDigitos = False: Exit Function
    Next i
    EsSoloDigitos = True
End Function

Private Function MapaArchivos(ByVal carpeta As String, ByVal patron As String, _
                              ByVal extOK As String) As Object
    Dim d As Object, f As String, fch As Date, ext As String
    Set d = CreateObject("Scripting.Dictionary")
    carpeta = ConBarra(carpeta)
    If Len(carpeta) = 0 Then Set MapaArchivos = d: Exit Function
    If Len(Dir(carpeta, vbDirectory)) = 0 Then Set MapaArchivos = d: Exit Function
    f = Dir(carpeta & patron)
    Do While Len(f) > 0
        ext = LCase$(Mid$(f, InStrRev(f, ".")))
        If ext = LCase$(extOK) Then
            fch = Fecha8(f)
            If fch > 0 Then
                If Not d.Exists(CLng(fch)) Then d.Add CLng(fch), carpeta & f
            End If
        End If
        f = Dir
    Loop
    Set MapaArchivos = d
End Function

Private Function MapaFMS() As Object
    Set MapaFMS = MapaArchivos(Cfg("B1"), "FMS_*.xls*", ".xlsx")
End Function

Private Function MapaVector() As Object
    Set MapaVector = MapaArchivos(Cfg("B2"), "*RFL.xls*", ".xls")
End Function

Private Function ClavesOrdenadas(d As Object) As Long()
    Dim arr() As Long, k As Variant, n As Long
    If d.Count = 0 Then ReDim arr(0 To 0): ClavesOrdenadas = arr: Exit Function
    ReDim arr(1 To d.Count)
    For Each k In d.Keys
        n = n + 1: arr(n) = CLng(k)
    Next k
    OrdenarLong arr
    ClavesOrdenadas = arr
End Function

Private Sub OrdenarLong(ByRef a() As Long)
    Dim i As Long, j As Long, t As Long
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) < a(i) Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
End Sub

Private Function FMSParaVector(ByVal v As Long, ByRef arrFMS() As Long, _
                               ByVal desfase As Long) As Long
    Dim i As Long, idx As Long
    idx = 0
    For i = LBound(arrFMS) To UBound(arrFMS)
        If arrFMS(i) <= v Then idx = i Else Exit For
    Next i
    If idx = 0 Then FMSParaVector = 0: Exit Function
    If desfase = 0 Then
        If arrFMS(idx) = v Then FMSParaVector = v Else FMSParaVector = 0
        Exit Function
    End If
    If arrFMS(idx) = v Then idx = idx - desfase Else idx = idx - (desfase - 1)
    If idx < LBound(arrFMS) Then FMSParaVector = 0 Else FMSParaVector = arrFMS(idx)
End Function

Private Function Observaciones() As Variant
    Dim dF As Object, dV As Object
    Dim arrF() As Long, arrV() As Long, res() As Long
    Dim i As Long, n As Long, desf As Long
    Dim fIni As Date, fFin As Date, fch As Date

    Set dF = MapaFMS(): Set dV = MapaVector()
    If dF.Count = 0 Or dV.Count = 0 Then Observaciones = Array(): Exit Function

    arrF = ClavesOrdenadas(dF): arrV = ClavesOrdenadas(dV)
    desf = CLng(CfgNum("B9", 1))
    fIni = CfgFecha("B4"): fFin = CfgFecha("B5")

    ReDim res(1 To UBound(arrV))
    For i = LBound(arrV) To UBound(arrV)
        If FMSParaVector(arrV(i), arrF, desf) > 0 Then
            fch = CDate(arrV(i))
            If (fIni = 0 Or fch >= fIni) And (fFin = 0 Or fch <= fFin) Then
                n = n + 1: res(n) = arrV(i)
            End If
        End If
    Next i
    If n = 0 Then Observaciones = Array(): Exit Function
    ReDim Preserve res(1 To n)
    Observaciones = res
End Function

'============================================================
' INVENTARIO
'============================================================
Public Sub Inventariar()
    Dim dF As Object, dV As Object, ws As Worksheet
    Dim arrF() As Long, arrV() As Long
    Dim i As Long, fila As Long, desf As Long
    Dim fV As Date, fFMS As Long, nOK As Long, nSin As Long

    Set dF = MapaFMS(): Set dV = MapaVector()
    If dF.Count = 0 Then MsgBox "No encontre FMS en: " & Cfg("B1"), vbExclamation: Exit Sub
    If dV.Count = 0 Then MsgBox "No encontre vectores en: " & Cfg("B2"), vbExclamation: Exit Sub

    arrF = ClavesOrdenadas(dF): arrV = ClavesOrdenadas(dV)
    desf = CLng(CfgNum("B9", 1))

    Set ws = Hoja(SH_INV)
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("FECHA OBS (VECTOR)", "DIA", "FMS PAREADO", "DIAS DESDE ANT", "ESTADO")

    fila = 2
    For i = LBound(arrV) To UBound(arrV)
        fV = CDate(arrV(i))
        fFMS = FMSParaVector(arrV(i), arrF, desf)
        ws.Cells(fila, 1).Value = fV
        ws.Cells(fila, 2).Value = Format(fV, "ddd")
        If fFMS > 0 Then
            ws.Cells(fila, 3).Value = CDate(fFMS)
            ws.Cells(fila, 5).Value = "OK": nOK = nOK + 1
        Else
            ws.Cells(fila, 3).Value = "-"
            ws.Cells(fila, 5).Value = "SIN FMS PAREABLE"
            ws.Cells(fila, 5).Font.Color = RGB(212, 12, 12)
            nSin = nSin + 1
        End If
        If fila > 2 Then ws.Cells(fila, 4).Value = fV - ws.Cells(fila - 1, 1).Value
        fila = fila + 1
    Next i

    ws.Range("A2:A" & fila - 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("C2:C" & fila - 1).NumberFormat = "dd/mm/yyyy"
    FormatoInstitucional ws.Range("A1:E" & fila - 1)
    ws.Columns("A:E").AutoFit

    MsgBox "Inventario (desfase " & desf & ")" & vbCrLf & _
           "Observaciones OK: " & nOK & vbCrLf & _
           "Sin FMS pareable: " & nSin & vbCrLf & _
           "Rango: " & Format(CDate(arrV(LBound(arrV))), "dd/mm/yyyy") & " a " & _
                       Format(CDate(arrV(UBound(arrV))), "dd/mm/yyyy") & vbCrLf & _
           "Col D: 1 y 3 es normal; >3 = falta archivo.", vbInformation
End Sub

'============================================================
' DEPOSITOS Y CURVA
'============================================================
Private Function CargarDepositos() As Object
    Dim d As Object, ws As Worksheet, uf As Long, i As Long, cod As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = Hoja(SH_DEP)
    uf = UltFila(ws)
    For i = 2 To uf
        cod = Trim$(CStr(ws.Cells(i, 3).Value))          ' C = codigo SBS
        If Len(cod) > 0 And IsNumeric(ws.Cells(i, 6).Value) Then  ' F = tasa
            If Not d.Exists(cod) Then d.Add cod, CDbl(ws.Cells(i, 6).Value)
        End If
    Next i
    Set CargarDepositos = d
End Function

' Interpola duracion desde CP Duracion (A=dias asc, B=duracion)
Private Function InterpolarDur(ByVal dias As Double) As Double
    Dim ws As Worksheet, uf As Long, i As Long
    Dim d0 As Double, d1 As Double, v0 As Double, v1 As Double
    Set ws = Hoja(SH_CUR)
    uf = UltFila(ws)
    InterpolarDur = 0
    If uf < 2 Then Exit Function
    If Not IsNumeric(ws.Cells(2, 1).Value) Then Exit Function

    If dias <= Val(ws.Cells(2, 1).Value) Then
        InterpolarDur = Val(ws.Cells(2, 2).Value): Exit Function
    End If
    If dias >= Val(ws.Cells(uf, 1).Value) Then
        InterpolarDur = Val(ws.Cells(uf, 2).Value): Exit Function
    End If
    For i = 2 To uf - 1
        d0 = Val(ws.Cells(i, 1).Value): d1 = Val(ws.Cells(i + 1, 1).Value)
        If dias >= d0 And dias <= d1 And d1 > d0 Then
            v0 = Val(ws.Cells(i, 2).Value): v1 = Val(ws.Cells(i + 1, 2).Value)
            InterpolarDur = v0 + (v1 - v0) * (dias - d0) / (d1 - d0)
            Exit Function
        End If
    Next i
End Function

'============================================================
' RATING (SUPUESTO: escala local LP + CP; ajustar si hace falta)
'============================================================
Private Function RatingScore(ByVal r As String) As Double
    Dim s As String
    s = UCase$(Replace(Trim$(r), " ", ""))
    Select Case True
        Case s Like "AAA*": RatingScore = 1
        Case s Like "AA+*": RatingScore = 2
        Case s Like "AA-*": RatingScore = 4
        Case s Like "AA*":  RatingScore = 3
        Case s Like "A+*":  RatingScore = 5
        Case s Like "A-*":  RatingScore = 7
        Case s Like "BBB+*": RatingScore = 8
        Case s Like "BBB-*": RatingScore = 10
        Case s Like "BBB*":  RatingScore = 9
        Case s Like "CP-1*", s Like "CP1*", s Like "CATEGORIAI*": RatingScore = 3
        Case s Like "CP-2*", s Like "CP2*": RatingScore = 6
        Case s Like "A*":   RatingScore = 6
        Case Else: RatingScore = 15        ' bajo IG o sin rating reconocido
    End Select
End Function

Private Function EsIG(ByVal r As String) As Boolean
    EsIG = (RatingScore(r) <= 10)
End Function

Private Function ScoreALetra(ByVal sc As Double) As String
    Select Case sc
        Case Is <= 1.5: ScoreALetra = "AAA"
        Case Is <= 2.5: ScoreALetra = "AA+"
        Case Is <= 3.5: ScoreALetra = "AA"
        Case Is <= 4.5: ScoreALetra = "AA-"
        Case Is <= 5.5: ScoreALetra = "A+"
        Case Is <= 6.5: ScoreALetra = "A"
        Case Is <= 7.5: ScoreALetra = "A-"
        Case Is <= 8.5: ScoreALetra = "BBB+"
        Case Is <= 9.5: ScoreALetra = "BBB"
        Case Is <= 10.5: ScoreALetra = "BBB-"
        Case Else: ScoreALetra = "<IG"
    End Select
End Function

'============================================================
' CATEGORIA
'============================================================
Public Function Categoria(ByVal assetClass As String) As String
    Dim s As String
    s = LCase$(Trim$(assetClass))
    s = Replace(s, "ó", "o"): s = Replace(s, "í", "i"): s = Replace(s, "é", "e")
    s = Replace(s, "á", "a"): s = Replace(s, "ú", "u")
    If InStr(s, "deposito") > 0 Then
        Categoria = "Depositos a Plazo"
    ElseIf InStr(s, "papel") > 0 Or InStr(s, "cd seriado") > 0 Then
        Categoria = "Papeles Comerciales"
    ElseIf InStr(s, "bono") > 0 Or InStr(s, "titulo") > 0 Then
        Categoria = "Bonos"
    Else
        Categoria = "Otros"
    End If
End Function

Private Function NormFond(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    Do While Left$(s, 1) = "0" And Len(s) > 1
        s = Mid$(s, 2)
    Loop
    NormFond = s
End Function

'============================================================
' MOTOR: CalcularDia
'   Lee vector + FMS pareado, cruza depositos, calcula metricas,
'   escribe fila en Historico (idempotente). Si escribirHojas,
'   regenera Cartera F0 y las hojas de metricas.
'============================================================
Public Function CalcularDia(ByVal fechaVector As Date, ByVal escribirHojas As Boolean) As Boolean
    Dim dF As Object, dV As Object, dVec As Object, dDep As Object
    Dim arrF() As Long, fFMS As Long, desf As Long
    Dim wbF As Workbook, wbV As Workbook, wsF As Worksheet
    Dim uf As Long, i As Long
    Dim colCruce As Long, codFondo As String
    Dim cod As String, monto As Double
    Dim cart() As Variant, nD As Long
    Dim v As Variant

    CalcularDia = False
    Set dF = MapaFMS(): Set dV = MapaVector()
    If Not dV.Exists(CLng(fechaVector)) Then Exit Function
    arrF = ClavesOrdenadas(dF)
    desf = CLng(CfgNum("B9", 1))
    fFMS = FMSParaVector(CLng(fechaVector), arrF, desf)
    If fFMS = 0 Then Exit Function

    colCruce = CLng(CfgNum("B8", 7))
    codFondo = Cfg("B7"): If Len(codFondo) = 0 Then codFondo = "9"
    Set dDep = CargarDepositos()

    On Error GoTo Fallo

    ' --- 1) Vector -> dict ---
    Set wbV = Workbooks.Open(dV(CLng(fechaVector)), ReadOnly:=True, UpdateLinks:=0)
    Set dVec = CreateObject("Scripting.Dictionary")
    With wbV.Worksheets(1)
        uf = UltFila(wbV.Worksheets(1))
        For i = 2 To uf
            cod = Trim$(CStr(.Cells(i, 1).Value))
            If Len(cod) > 0 Then
                If Not dVec.Exists(cod) Then
                    dVec.Add cod, Array(Val(.Cells(i, 14).Value), _
                                        Val(.Cells(i, 15).Value) * 100, _
                                        Val(.Cells(i, 24).Value))
                End If
            End If
        Next i
    End With
    wbV.Close SaveChanges:=False: Set wbV = Nothing

    ' --- 2) FMS -> matriz de cartera en memoria ---
    Set wbF = Workbooks.Open(dF(fFMS), ReadOnly:=True, UpdateLinks:=0)
    Set wsF = wbF.Worksheets("Cartera")
    uf = UltFila(wsF)
    ReDim cart(1 To uf, 1 To C_NCOLS)
    nD = 0

    For i = 2 To uf
        If NormFond(wsF.Cells(i, 1).Value) = codFondo Then
            nD = nD + 1
            cod = Trim$(CStr(wsF.Cells(i, colCruce).Value))
            monto = Val(wsF.Cells(i, 10).Value) / 1000000#

            cart(nD, C_COD) = cod
            cart(nD, C_EMI) = Trim$(CStr(wsF.Cells(i, 5).Value))
            cart(nD, C_AC) = Trim$(CStr(wsF.Cells(i, 3).Value))
            cart(nD, C_CAT) = Categoria(CStr(wsF.Cells(i, 3).Value))
            cart(nD, C_MON) = Trim$(CStr(wsF.Cells(i, 8).Value))
            cart(nD, C_CANT) = Val(wsF.Cells(i, 9).Value)
            cart(nD, C_MTO) = monto
            cart(nD, C_RAT) = Trim$(CStr(wsF.Cells(i, 11).Value))
            If IsDate(wsF.Cells(i, 12).Value) Then
                cart(nD, C_VCTO) = CDate(wsF.Cells(i, 12).Value)
                cart(nD, C_DIAS) = CDbl(CDate(wsF.Cells(i, 12).Value) - fechaVector)
            Else
                cart(nD, C_VCTO) = Empty: cart(nD, C_DIAS) = 0
            End If
            cart(nD, C_DEP) = IIf(cart(nD, C_CAT) = "Depositos a Plazo", 1, 0)

            If dVec.Exists(cod) Then
                v = dVec(cod)
                cart(nD, C_YTW) = v(0): cart(nD, C_SPR) = v(1): cart(nD, C_DUR) = v(2)
                cart(nD, C_VEC) = 1
            ElseIf cart(nD, C_DEP) = 1 Then
                ' deposito: tasa pactada + duracion interpolada
                If dDep.Exists(cod) Then cart(nD, C_YTW) = dDep(cod) Else cart(nD, C_YTW) = 0
                cart(nD, C_SPR) = 0
                cart(nD, C_DUR) = InterpolarDur(cart(nD, C_DIAS))
                cart(nD, C_VEC) = 0
            Else
                cart(nD, C_YTW) = 0: cart(nD, C_SPR) = 0: cart(nD, C_DUR) = 0
                cart(nD, C_VEC) = 0
            End If
        End If
    Next i
    wbF.Close SaveChanges:=False: Set wbF = Nothing

    If nD = 0 Then Exit Function

    ' --- 3) Metricas + Historico (+ hojas si aplica) ---
    ProcesarDia fechaVector, CDate(fFMS), cart, nD, escribirHojas
    CalcularDia = True
    Exit Function

Fallo:
    On Error Resume Next
    If Not wbV Is Nothing Then wbV.Close SaveChanges:=False
    If Not wbF Is Nothing Then wbF.Close SaveChanges:=False
End Function

'============================================================
' PROCESAR: metricas del dia + fila de Historico + hojas
'============================================================
Private Sub ProcesarDia(ByVal fObs As Date, ByVal fFMS As Date, _
                        ByRef cart() As Variant, ByVal nD As Long, _
                        ByVal escribirHojas As Boolean)
    Dim i As Long
    Dim mtoTot As Double, mtoNeg As Double, mtoDep As Double
    Dim mtoBon As Double, mtoPC As Double
    Dim ytmF As Double, ytmN As Double, ytmB As Double, ytmP As Double, ytmD As Double
    Dim mtoYtmF As Double, mtoYtmN As Double   ' montos con tasa valida
    Dim sprT As Double, sprB As Double, sprP As Double
    Dim durT As Double, durB As Double, durP As Double, durD As Double
    Dim mtoIG As Double, mtoRated As Double, scoreAcum As Double
    Dim v30 As Double, v90 As Double
    Dim nSin As Long
    Dim dEmis As Object, k As Variant
    Dim carryDia As Double, diasTr As Double

    Set dEmis = CreateObject("Scripting.Dictionary")

    For i = 1 To nD
        mtoTot = mtoTot + cart(i, C_MTO)

        If cart(i, C_DEP) = 1 Then
            mtoDep = mtoDep + cart(i, C_MTO)
            If cart(i, C_YTW) > 0 Then ytmD = ytmD + cart(i, C_MTO) * cart(i, C_YTW)
            durD = durD + cart(i, C_MTO) * cart(i, C_DUR)
        Else
            mtoNeg = mtoNeg + cart(i, C_MTO)
            If cart(i, C_CAT) = "Bonos" Then
                mtoBon = mtoBon + cart(i, C_MTO)
                ytmB = ytmB + cart(i, C_MTO) * cart(i, C_YTW)
                sprB = sprB + cart(i, C_MTO) * cart(i, C_SPR)
                durB = durB + cart(i, C_MTO) * cart(i, C_DUR)
            ElseIf cart(i, C_CAT) = "Papeles Comerciales" Then
                mtoPC = mtoPC + cart(i, C_MTO)
                ytmP = ytmP + cart(i, C_MTO) * cart(i, C_YTW)
                sprP = sprP + cart(i, C_MTO) * cart(i, C_SPR)
                durP = durP + cart(i, C_MTO) * cart(i, C_DUR)
            End If
            If cart(i, C_VEC) = 1 Then
                ytmN = ytmN + cart(i, C_MTO) * cart(i, C_YTW)
                mtoYtmN = mtoYtmN + cart(i, C_MTO)
                sprT = sprT + cart(i, C_MTO) * cart(i, C_SPR)
            Else
                nSin = nSin + 1
            End If
            ' rating solo universo negociable
            If Len(CStr(cart(i, C_RAT))) > 0 Then
                mtoRated = mtoRated + cart(i, C_MTO)
                scoreAcum = scoreAcum + cart(i, C_MTO) * RatingScore(CStr(cart(i, C_RAT)))
                If EsIG(CStr(cart(i, C_RAT))) Then mtoIG = mtoIG + cart(i, C_MTO)
            End If
            ' concentracion por emisor (negociable)
            k = cart(i, C_EMI)
            If dEmis.Exists(k) Then dEmis(k) = dEmis(k) + cart(i, C_MTO) Else dEmis.Add k, cart(i, C_MTO)
        End If

        durT = durT + cart(i, C_MTO) * cart(i, C_DUR)
        If cart(i, C_DIAS) > 0 And cart(i, C_DIAS) <= 30 Then v30 = v30 + cart(i, C_MTO)
        If cart(i, C_DIAS) > 0 And cart(i, C_DIAS) <= 90 Then v90 = v90 + cart(i, C_MTO)

        ' YTM fondo: todo lo que tenga tasa (vector o deposito con tasa)
        If cart(i, C_YTW) > 0 Then
            ytmF = ytmF + cart(i, C_MTO) * cart(i, C_YTW)
            mtoYtmF = mtoYtmF + cart(i, C_MTO)
        End If
    Next i

    If mtoYtmF > 0 Then ytmF = ytmF / mtoYtmF
    If mtoYtmN > 0 Then ytmN = ytmN / mtoYtmN
    If mtoBon > 0 Then ytmB = ytmB / mtoBon: sprB = sprB / mtoBon: durB = durB / mtoBon
    If mtoPC > 0 Then ytmP = ytmP / mtoPC: sprP = sprP / mtoPC: durP = durP / mtoPC
    If mtoDep > 0 Then ytmD = ytmD / mtoDep: durD = durD / mtoDep
    If mtoYtmN > 0 Then sprT = sprT / mtoYtmN
    If mtoTot > 0 Then durT = durT / mtoTot
    If mtoRated > 0 Then scoreAcum = scoreAcum / mtoRated

    ' HHI y TopN por emisor (universo negociable, pesos renormalizados)
    Dim hhi As Double, pesos() As Double, nE As Long, t5 As Double, t10 As Double, w As Double
    nE = dEmis.Count
    If nE > 0 And mtoNeg > 0 Then
        ReDim pesos(1 To nE)
        i = 0
        For Each k In dEmis.Keys
            i = i + 1
            w = dEmis(k) / mtoNeg
            pesos(i) = w
            hhi = hhi + (w * 100) ^ 2
        Next k
        OrdenarDblDesc pesos
        For i = 1 To WorksheetFunction.Min(5, nE): t5 = t5 + pesos(i): Next i
        For i = 1 To WorksheetFunction.Min(10, nE): t10 = t10 + pesos(i): Next i
    End If

    ' dias transcurridos desde la observacion anterior en Historico
    diasTr = DiasDesdeAnterior(fObs)
    ' carry del dia (%): YTM fondo anualizado a los dias transcurridos
    carryDia = ytmF / CfgNum("B6", 365) * diasTr

    ' --- fila de Historico (idempotente) ---
    EscribirHistorico fObs, diasTr, mtoTot, ytmF, ytmN, ytmB, ytmP, ytmD, _
                      sprT, sprB, sprP, durT, durB, durP, durD, _
                      IIf(mtoRated > 0, mtoIG / mtoRated, 0), ScoreALetra(scoreAcum), _
                      hhi, t5, t10, v30, v90, carryDia, nD, nSin

    ' --- hojas del dia ---
    If escribirHojas Then
        EscribirCartera fObs, fFMS, cart, nD, mtoTot
        EscribirMetricas fObs, cart, nD, mtoTot, mtoNeg, dEmis, diasTr
    End If
End Sub

Private Sub OrdenarDblDesc(ByRef a() As Double)
    Dim i As Long, j As Long, t As Double
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) > a(i) Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
End Sub

Private Function DiasDesdeAnterior(ByVal fObs As Date) As Double
    Dim ws As Worksheet, uf As Long, i As Long
    Dim mejor As Date, f As Variant
    Set ws = Hoja(SH_HIS)
    uf = UltFila(ws)
    mejor = 0
    For i = 2 To uf
        f = ws.Cells(i, 1).Value
        If IsDate(f) Then
            If CDate(f) < fObs And CDate(f) > mejor Then mejor = CDate(f)
        End If
    Next i
    If mejor = 0 Then DiasDesdeAnterior = 1 Else DiasDesdeAnterior = CDbl(fObs - mejor)
End Function

Private Sub EscribirHistorico(ByVal fObs As Date, ByVal diasTr As Double, _
    ByVal valCart As Double, ByVal ytmF As Double, ByVal ytmN As Double, _
    ByVal ytmB As Double, ByVal ytmP As Double, ByVal ytmD As Double, _
    ByVal sprT As Double, ByVal sprB As Double, ByVal sprP As Double, _
    ByVal durT As Double, ByVal durB As Double, ByVal durP As Double, ByVal durD As Double, _
    ByVal pctIG As Double, ByVal ratProm As String, _
    ByVal hhi As Double, ByVal t5 As Double, ByVal t10 As Double, _
    ByVal v30 As Double, ByVal v90 As Double, ByVal carryDia As Double, _
    ByVal nInstr As Long, ByVal nSin As Long)

    Dim ws As Worksheet, uf As Long, i As Long, fila As Long
    Set ws = Hoja(SH_HIS)
    If UltFila(ws) = 0 Then EncabezadoHistorico ws
    uf = UltFila(ws)

    ' idempotente: buscar fecha
    fila = 0
    For i = 2 To uf
        If IsDate(ws.Cells(i, 1).Value) Then
            If CDate(ws.Cells(i, 1).Value) = fObs Then fila = i: Exit For
        End If
    Next i
    If fila = 0 Then fila = uf + 1

    ws.Cells(fila, 1).Value = fObs
    ws.Cells(fila, 2).Value = diasTr
    ws.Cells(fila, 3).Value = valCart
    ws.Cells(fila, 4).Value = ytmF
    ws.Cells(fila, 5).Value = ytmN
    ws.Cells(fila, 6).Value = ytmB
    ws.Cells(fila, 7).Value = ytmP
    ws.Cells(fila, 8).Value = ytmD
    ws.Cells(fila, 9).Value = sprT
    ws.Cells(fila, 10).Value = sprB
    ws.Cells(fila, 11).Value = sprP
    ws.Cells(fila, 12).Value = durT
    ws.Cells(fila, 13).Value = durB
    ws.Cells(fila, 14).Value = durP
    ws.Cells(fila, 15).Value = durD
    ws.Cells(fila, 16).Value = pctIG
    ws.Cells(fila, 17).Value = ratProm
    ws.Cells(fila, 18).Value = hhi
    ws.Cells(fila, 19).Value = t5
    ws.Cells(fila, 20).Value = t10
    ws.Cells(fila, 21).Value = v30
    ws.Cells(fila, 22).Value = v90
    ws.Cells(fila, 23).Value = carryDia

    ws.Cells(fila, 26).Value = nInstr
    ws.Cells(fila, 27).Value = nSin

    ' ordenar por fecha y recalcular carry acumulado MTD/YTD de toda la hoja
    uf = UltFila(ws)
    If uf > 2 Then
        ws.Range("A2:AA" & uf).Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlNo
    End If
    RecalcularCarryAcum ws

    ws.Range("A2:A" & uf).NumberFormat = "dd/mm/yyyy"
    ws.Range("C2:C" & uf).NumberFormat = "#,##0.00"
    ws.Range("D2:H" & uf).NumberFormat = "0.0000"
    ws.Range("I2:K" & uf).NumberFormat = "#,##0.0"
    ws.Range("L2:O" & uf).NumberFormat = "0.00"
    ws.Range("P2:P" & uf).NumberFormat = "0.0%"
    ws.Range("R2:R" & uf).NumberFormat = "#,##0"
    ws.Range("S2:T" & uf).NumberFormat = "0.0%"
    ws.Range("U2:V" & uf).NumberFormat = "#,##0.00"
    ws.Range("W2:Y" & uf).NumberFormat = "0.0000%"
End Sub

Private Sub RecalcularCarryAcum(ws As Worksheet)
    Dim uf As Long, i As Long
    Dim f As Date, mtd As Double, ytd As Double
    Dim mesAct As Long, anioAct As Long
    uf = UltFila(ws)
    mesAct = -1: anioAct = -1
    For i = 2 To uf
        If IsDate(ws.Cells(i, 1).Value) Then
            f = CDate(ws.Cells(i, 1).Value)
            If Year(f) <> anioAct Then anioAct = Year(f): ytd = 0: mesAct = -1
            If Month(f) <> mesAct Then mesAct = Month(f): mtd = 0
            mtd = mtd + Val(ws.Cells(i, 23).Value) / 100
            ytd = ytd + Val(ws.Cells(i, 23).Value) / 100
            ws.Cells(i, 24).Value = mtd
            ws.Cells(i, 25).Value = ytd
        End If
    Next i
End Sub

'============================================================
' HOJAS DEL DIA
'============================================================
Private Sub EscribirCartera(ByVal fObs As Date, ByVal fFMS As Date, _
                            ByRef cart() As Variant, ByVal nD As Long, ByVal mtoTot As Double)
    Dim ws As Worksheet, i As Long, fila As Long
    Set ws = Hoja(SH_CART)
    ws.Cells.Clear
    ws.Range("A1:R1").Value = Array("FECHA OBS", "FECHA FMS", "CODIGO", "EMISOR", "ASSET CLASS", _
        "CATEGORIA", "MON", "CANTIDAD", "MONTO MM", "RATING", "VCTO", "DIAS", _
        "YTW", "SPREAD PBS", "DURACION", "EN VECTOR", "DEPOSITO", "PESO")
    For i = 1 To nD
        ws.Cells(i + 1, 1).Value = fObs
        ws.Cells(i + 1, 2).Value = fFMS
        ws.Cells(i + 1, 3).Value = cart(i, C_COD)
        ws.Cells(i + 1, 4).Value = cart(i, C_EMI)
        ws.Cells(i + 1, 5).Value = cart(i, C_AC)
        ws.Cells(i + 1, 6).Value = cart(i, C_CAT)
        ws.Cells(i + 1, 7).Value = cart(i, C_MON)
        ws.Cells(i + 1, 8).Value = cart(i, C_CANT)
        ws.Cells(i + 1, 9).Value = cart(i, C_MTO)
        ws.Cells(i + 1, 10).Value = cart(i, C_RAT)
        If Not IsEmpty(cart(i, C_VCTO)) Then ws.Cells(i + 1, 11).Value = cart(i, C_VCTO)
        ws.Cells(i + 1, 12).Value = cart(i, C_DIAS)
        ws.Cells(i + 1, 13).Value = cart(i, C_YTW)
        ws.Cells(i + 1, 14).Value = cart(i, C_SPR)
        ws.Cells(i + 1, 15).Value = cart(i, C_DUR)
        ws.Cells(i + 1, 16).Value = IIf(cart(i, C_VEC) = 1, "SI", "NO")
        ws.Cells(i + 1, 17).Value = IIf(cart(i, C_DEP) = 1, "SI", "NO")
        If mtoTot > 0 Then ws.Cells(i + 1, 18).Value = cart(i, C_MTO) / mtoTot
    Next i
    fila = nD + 1
    ws.Range("A2:B" & fila).NumberFormat = "dd/mm/yyyy"
    ws.Range("K2:K" & fila).NumberFormat = "dd/mm/yyyy"
    ws.Range("H2:I" & fila).NumberFormat = "#,##0.00"
    ws.Range("M2:M" & fila).NumberFormat = "0.0000"
    ws.Range("N2:O" & fila).NumberFormat = "#,##0.00"
    ws.Range("R2:R" & fila).NumberFormat = "0.00%"
    FormatoInstitucional ws.Range("A1:R" & fila)
    ws.Columns("A:R").AutoFit
End Sub

Private Sub EscribirMetricas(ByVal fObs As Date, ByRef cart() As Variant, ByVal nD As Long, _
                             ByVal mtoTot As Double, ByVal mtoNeg As Double, _
                             dEmis As Object, ByVal diasTr As Double)
    Dim ws As Worksheet, i As Long, fila As Long, k As Variant

    ' ---------- YIELD: por categoria y por tramo ----------
    Dim cats As Variant, tramos As Variant
    cats = Array("Bonos", "Papeles Comerciales", "Depositos a Plazo")
    Dim mtoC As Double, ytmC As Double

    Set ws = Hoja(SH_YLD)
    ws.Cells.Clear
    ws.Range("A1").Value = "YIELD — " & Format(fObs, "dd/mm/yyyy")
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:C3").Value = Array("CATEGORIA", "MONTO MM", "YTM POND")
    fila = 4
    For i = 0 To 2
        mtoC = 0: ytmC = 0
        Dim j As Long
        For j = 1 To nD
            If cart(j, C_CAT) = cats(i) And cart(j, C_YTW) > 0 Then
                mtoC = mtoC + cart(j, C_MTO)
                ytmC = ytmC + cart(j, C_MTO) * cart(j, C_YTW)
            End If
        Next j
        ws.Cells(fila, 1).Value = cats(i)
        ws.Cells(fila, 2).Value = mtoC
        If mtoC > 0 Then ws.Cells(fila, 3).Value = ytmC / mtoC
        fila = fila + 1
    Next i
    FormatoInstitucional ws.Range("A3:C" & fila - 1)

    ' por tramo (universo negociable)
    ws.Cells(fila + 1, 1).Value = "POR TRAMO (negociable)"
    ws.Cells(fila + 1, 1).Font.Bold = True
    ws.Range(ws.Cells(fila + 2, 1), ws.Cells(fila + 2, 3)).Value = Array("TRAMO", "MONTO MM", "YTM POND")
    Dim r0 As Long: r0 = fila + 2
    tramos = Array(Array("0-90d", 0, 90), Array("90-180d", 90, 180), _
                   Array("180-360d", 180, 360), Array(">360d", 360, 100000))
    fila = r0 + 1
    For i = 0 To 3
        mtoC = 0: ytmC = 0
        For j = 1 To nD
            If cart(j, C_DEP) = 0 And cart(j, C_VEC) = 1 Then
                If cart(j, C_DIAS) > tramos(i)(1) And cart(j, C_DIAS) <= tramos(i)(2) Then
                    mtoC = mtoC + cart(j, C_MTO)
                    ytmC = ytmC + cart(j, C_MTO) * cart(j, C_YTW)
                End If
            End If
        Next j
        ws.Cells(fila, 1).Value = tramos(i)(0)
        ws.Cells(fila, 2).Value = mtoC
        If mtoC > 0 Then ws.Cells(fila, 3).Value = ytmC / mtoC
        fila = fila + 1
    Next i
    FormatoInstitucional ws.Range(ws.Cells(r0, 1), ws.Cells(fila - 1, 3))
    ws.Columns("A:C").AutoFit
    ws.Range("B4:B" & fila).NumberFormat = "#,##0.00"
    ws.Range("C4:C" & fila).NumberFormat = "0.0000"

    ' ---------- CALIDAD: distribucion por rating (negociable) ----------
    Dim dRat As Object
    Set dRat = CreateObject("Scripting.Dictionary")
    For i = 1 To nD
        If cart(i, C_DEP) = 0 And Len(CStr(cart(i, C_RAT))) > 0 Then
            k = cart(i, C_RAT)
            If dRat.Exists(k) Then dRat(k) = dRat(k) + cart(i, C_MTO) Else dRat.Add k, cart(i, C_MTO)
        End If
    Next i
    Set ws = Hoja(SH_CAL)
    ws.Cells.Clear
    ws.Range("A1").Value = "CALIDAD CREDITICIA (negociable) — " & Format(fObs, "dd/mm/yyyy")
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:C3").Value = Array("RATING", "MONTO MM", "PESO")
    fila = 4
    For Each k In dRat.Keys
        ws.Cells(fila, 1).Value = k
        ws.Cells(fila, 2).Value = dRat(k)
        If mtoNeg > 0 Then ws.Cells(fila, 3).Value = dRat(k) / mtoNeg
        fila = fila + 1
    Next k
    If fila > 4 Then
        ws.Range("A4:C" & fila - 1).Sort Key1:=ws.Range("B4"), Order1:=xlDescending, Header:=xlNo
        FormatoInstitucional ws.Range("A3:C" & fila - 1)
    End If
    ws.Range("B4:B" & fila).NumberFormat = "#,##0.00"
    ws.Range("C4:C" & fila).NumberFormat = "0.00%"
    ws.Columns("A:C").AutoFit

    ' ---------- CONCENTRACION: top emisores (negociable) ----------
    Set ws = Hoja(SH_CON)
    ws.Cells.Clear
    ws.Range("A1").Value = "CONCENTRACION POR EMISOR (negociable) — " & Format(fObs, "dd/mm/yyyy")
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:C3").Value = Array("EMISOR", "MONTO MM", "PESO")
    fila = 4
    For Each k In dEmis.Keys
        ws.Cells(fila, 1).Value = k
        ws.Cells(fila, 2).Value = dEmis(k)
        If mtoNeg > 0 Then ws.Cells(fila, 3).Value = dEmis(k) / mtoNeg
        fila = fila + 1
    Next k
    If fila > 4 Then
        ws.Range("A4:C" & fila - 1).Sort Key1:=ws.Range("B4"), Order1:=xlDescending, Header:=xlNo
        FormatoInstitucional ws.Range("A3:C" & fila - 1)
    End If
    ws.Range("B4:B" & fila).NumberFormat = "#,##0.00"
    ws.Range("C4:C" & fila).NumberFormat = "0.00%"
    ws.Columns("A:C").AutoFit

    ' ---------- VENCIMIENTOS: año x categoria ----------
    Dim dAnio As Object
    Set dAnio = CreateObject("Scripting.Dictionary")
    For i = 1 To nD
        If Not IsEmpty(cart(i, C_VCTO)) Then
            k = Year(CDate(cart(i, C_VCTO))) & "|" & cart(i, C_CAT)
            If dAnio.Exists(k) Then dAnio(k) = dAnio(k) + cart(i, C_MTO) Else dAnio.Add k, cart(i, C_MTO)
        End If
    Next i
    Set ws = Hoja(SH_VEN)
    ws.Cells.Clear
    ws.Range("A1").Value = "VENCIMIENTOS (S/ MM) — " & Format(fObs, "dd/mm/yyyy")
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:E3").Value = Array("AÑO", "BONOS", "DEP PLAZO", "PAPELES COM", "TOTAL")
    Dim anios As Object, arrA() As Long, nA As Long
    Set anios = CreateObject("Scripting.Dictionary")
    For i = 1 To nD
        If Not IsEmpty(cart(i, C_VCTO)) Then
            If Not anios.Exists(CLng(Year(CDate(cart(i, C_VCTO))))) Then _
                anios.Add CLng(Year(CDate(cart(i, C_VCTO)))), 1
        End If
    Next i
    arrA = ClavesOrdenadas(anios)
    fila = 4
    For i = LBound(arrA) To UBound(arrA)
        If arrA(i) > 0 Then
            ws.Cells(fila, 1).Value = arrA(i)
            ws.Cells(fila, 2).Value = ValorDicc(dAnio, arrA(i) & "|Bonos")
            ws.Cells(fila, 3).Value = ValorDicc(dAnio, arrA(i) & "|Depositos a Plazo")
            ws.Cells(fila, 4).Value = ValorDicc(dAnio, arrA(i) & "|Papeles Comerciales")
            ws.Cells(fila, 5).Value = Val(ws.Cells(fila, 2).Value) + Val(ws.Cells(fila, 3).Value) + Val(ws.Cells(fila, 4).Value)
            fila = fila + 1
        End If
    Next i
    FormatoInstitucional ws.Range("A3:E" & fila - 1)
    ws.Range("B4:E" & fila).NumberFormat = "#,##0.00"
    ws.Columns("A:E").AutoFit

    ' ---------- CONTRIBUCION: carry por instrumento (negociable) + deps agregado ----------
    Set ws = Hoja(SH_CTR)
    ws.Cells.Clear
    ws.Range("A1").Value = "CONTRIBUCION CARRY (" & diasTr & " dias) — " & Format(fObs, "dd/mm/yyyy")
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:F3").Value = Array("CODIGO", "EMISOR", "CATEGORIA", "PESO FONDO", "CARRY %", "CONTRIB %")
    fila = 4
    Dim base As Double, carryI As Double, ctrDep As Double, pesoDep As Double
    base = CfgNum("B6", 365)
    For i = 1 To nD
        If cart(i, C_YTW) > 0 And mtoTot > 0 Then
            carryI = cart(i, C_YTW) / base * diasTr
            If cart(i, C_DEP) = 1 Then
                ctrDep = ctrDep + (cart(i, C_MTO) / mtoTot) * carryI
                pesoDep = pesoDep + cart(i, C_MTO) / mtoTot
            Else
                ws.Cells(fila, 1).Value = cart(i, C_COD)
                ws.Cells(fila, 2).Value = cart(i, C_EMI)
                ws.Cells(fila, 3).Value = cart(i, C_CAT)
                ws.Cells(fila, 4).Value = cart(i, C_MTO) / mtoTot
                ws.Cells(fila, 5).Value = carryI / 100
                ws.Cells(fila, 6).Value = (cart(i, C_MTO) / mtoTot) * carryI / 100
                fila = fila + 1
            End If
        End If
    Next i
    ' linea agregada de depositos
    ws.Cells(fila, 1).Value = "-"
    ws.Cells(fila, 2).Value = "DEPOSITOS A PLAZO (agregado)"
    ws.Cells(fila, 3).Value = "Depositos a Plazo"
    ws.Cells(fila, 4).Value = pesoDep
    ws.Cells(fila, 6).Value = ctrDep / 100
    fila = fila + 1
    If fila > 5 Then
        ws.Range("A4:F" & fila - 1).Sort Key1:=ws.Range("F4"), Order1:=xlDescending, Header:=xlNo
    End If
    FormatoInstitucional ws.Range("A3:F" & fila - 1)
    ws.Range("D4:D" & fila).NumberFormat = "0.00%"
    ws.Range("E4:F" & fila).NumberFormat = "0.0000%"
    ws.Columns("A:F").AutoFit
End Sub

Private Function ValorDicc(d As Object, ByVal k As String) As Double
    If d.Exists(k) Then ValorDicc = d(k) Else ValorDicc = 0
End Function

'============================================================
' BOTONES
'============================================================
Public Sub CorrerDiario()
    Dim o As Variant
    o = Observaciones()
    If Not IsArray(o) Then MsgBox "Sin observaciones. Revisa Config y corre Inventariar.", vbExclamation: Exit Sub
    If UBound(o) < 1 Then MsgBox "Sin observaciones pareables.", vbExclamation: Exit Sub

    PrepararApp True
    If CalcularDia(CDate(o(UBound(o))), True) Then
        PrepararApp False
        MsgBox "Dia calculado: " & Format(CDate(o(UBound(o))), "dd/mm/yyyy") & vbCrLf & _
               "Revisa Cartera F0 (columna EN VECTOR) e Historico.", vbInformation
    Else
        PrepararApp False
        MsgBox "Fallo el calculo. Revisa rutas y archivos.", vbExclamation
    End If
End Sub

Public Sub CorrerBackfill()
    Dim o As Variant, i As Long, nHech As Long, nSalt As Long, nFall As Long
    Dim wsH As Worksheet, fechasHechas As Object, uf As Long
    Dim esUltima As Boolean

    o = Observaciones()
    If Not IsArray(o) Then MsgBox "Sin observaciones.", vbExclamation: Exit Sub
    If UBound(o) < 1 Then MsgBox "Sin observaciones pareables.", vbExclamation: Exit Sub

    ' fechas ya calculadas (reanudable)
    Set fechasHechas = CreateObject("Scripting.Dictionary")
    Set wsH = Hoja(SH_HIS)
    uf = UltFila(wsH)
    For i = 2 To uf
        If IsDate(wsH.Cells(i, 1).Value) Then
            If Not fechasHechas.Exists(CLng(CDate(wsH.Cells(i, 1).Value))) Then _
                fechasHechas.Add CLng(CDate(wsH.Cells(i, 1).Value)), 1
        End If
    Next i

    If MsgBox("Backfill de " & (UBound(o) - LBound(o) + 1) & " observaciones" & vbCrLf & _
              "(ya calculadas: " & fechasHechas.Count & ", se saltan)." & vbCrLf & _
              "Puede tardar. ¿Continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    PrepararApp True
    For i = LBound(o) To UBound(o)
        esUltima = (i = UBound(o))
        If fechasHechas.Exists(CLng(o(i))) And Not esUltima Then
            nSalt = nSalt + 1
        Else
            Application.StatusBar = "Backfill " & (i - LBound(o) + 1) & "/" & _
                (UBound(o) - LBound(o) + 1) & "  " & Format(CDate(o(i)), "dd/mm/yyyy")
            If CalcularDia(CDate(o(i)), esUltima) Then nHech = nHech + 1 Else nFall = nFall + 1
        End If
    Next i
    PrepararApp False

    MsgBox "Backfill terminado." & vbCrLf & _
           "Calculadas: " & nHech & vbCrLf & _
           "Saltadas (ya estaban): " & nSalt & vbCrLf & _
           "Fallidas: " & nFall & IIf(nFall > 0, "  <- revisar archivos de esas fechas", ""), vbInformation
End Sub

Private Sub PrepararApp(ByVal activar As Boolean)
    Application.ScreenUpdating = Not activar
    Application.DisplayAlerts = Not activar
    Application.Calculation = IIf(activar, xlCalculationManual, xlCalculationAutomatic)
    If Not activar Then Application.StatusBar = False
End Sub

'============================================================
' FORMATO DE TABLAS
'============================================================
Public Sub FormatoInstitucional(rg As Range)
    With rg
        .Font.Name = "Arial"
        .Font.Size = 8
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    With rg.Rows(1)
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With rg.Rows(rg.Rows.Count)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
End Sub
