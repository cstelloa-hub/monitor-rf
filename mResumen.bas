Attribute VB_Name = "mResumen"
Option Explicit

'============================================================
' MONITOR FONDO 0 - Hoja Resumen (modulo aparte, v1)
' Lee Historico + hojas del dia y arma el tablero:
'   - KPIs con delta vs observacion anterior
'   - Semaforo de duracion (<1)
'   - Carry devengado MTD / YTD
'   - Top 5 emisores + calidad crediticia
'   - Grafico 1: YTM fondo (eje izq) + spread (eje der), serie diaria
'   - Grafico 2: vencimientos por año apilados por categoria
'
' MACRO: GenerarResumen  (correr despues de CorrerDiario)
'============================================================

Private Const SH_RES As String = "Resumen"
Private Const SH_HIS As String = "Historico"
Private Const SH_CON As String = "Concentracion"
Private Const SH_CAL As String = "Calidad"
Private Const SH_VEN As String = "Vencimientos"
Private Const SH_CART As String = "Cartera F0"

Private Const ROJO As Long = 786644        ' RGB(212,12,12)
Private Const GRIS As Long = 5855577       ' RGB(89,89,89)
Private Const ARENA As Long = 8431807      ' RGB(191,168,128)

Public Sub GenerarResumen()
    Dim wsH As Worksheet, ws As Worksheet
    Dim ufH As Long, fila As Long
    Dim tieneAnt As Boolean

    Set wsH = HojaR(SH_HIS)
    ufH = UltFilaR(wsH)
    If ufH < 2 Then
        MsgBox "El Historico esta vacio. Corre CorrerDiario o CorrerBackfill primero.", vbExclamation
        Exit Sub
    End If
    tieneAnt = (ufH >= 3)

    Application.ScreenUpdating = False

    Set ws = HojaR(SH_RES)
    ws.Cells.Clear
    BorrarGraficos ws

    ' ---------- Titulo ----------
    ws.Range("A1").Value = "MONITOR FONDO 0 — RESUMEN"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = "Observacion: " & Format(wsH.Cells(ufH, 1).Value, "dd/mm/yyyy") & _
                           "   |   Instrumentos: " & wsH.Cells(ufH, 26).Value & _
                           "   |   Generado: " & Format(Now, "dd/mm/yyyy hh:mm")
    ws.Range("A2").Font.Color = RGB(89, 89, 89)

    ' ---------- KPIs ----------
    ws.Range("A4").Value = "INDICADOR": ws.Range("B4").Value = "VALOR": ws.Range("C4").Value = "VS ANTERIOR"

    Dim vAct As Double, vAnt As Double

    ' Valor cartera
    ws.Range("A5").Value = "Valor cartera (MM)"
    vAct = Val(wsH.Cells(ufH, 3).Value)
    ws.Range("B5").Value = vAct
    ws.Range("B5").NumberFormat = "#,##0.00"
    If tieneAnt Then
        vAnt = Val(wsH.Cells(ufH - 1, 3).Value)
        ws.Range("C5").Value = vAct - vAnt
        ws.Range("C5").NumberFormat = "+#,##0.00;-#,##0.00"
    End If

    ' Duracion con semaforo
    ws.Range("A6").Value = "Duracion (limite < 1)"
    vAct = Val(wsH.Cells(ufH, 12).Value)
    ws.Range("B6").Value = vAct
    ws.Range("B6").NumberFormat = "0.0000"
    ws.Range("B6").Font.Bold = True
    If vAct >= 1 Then
        ws.Range("B6").Interior.Color = RGB(255, 150, 150)
        ws.Range("D6").Value = "EXCEDE LIMITE"
        ws.Range("D6").Font.Color = RGB(192, 0, 0): ws.Range("D6").Font.Bold = True
    Else
        ws.Range("B6").Interior.Color = RGB(180, 230, 180)
    End If
    If tieneAnt Then
        vAnt = Val(wsH.Cells(ufH - 1, 12).Value)
        ws.Range("C6").Value = vAct - vAnt
        ws.Range("C6").NumberFormat = "+0.0000;-0.0000"
    End If

    ' YTM fondo
    ws.Range("A7").Value = "YTM fondo (%)"
    vAct = Val(wsH.Cells(ufH, 4).Value)
    ws.Range("B7").Value = vAct
    ws.Range("B7").NumberFormat = "0.0000"
    If tieneAnt Then
        vAnt = Val(wsH.Cells(ufH - 1, 4).Value)
        ws.Range("C7").Value = (vAct - vAnt) * 100
        ws.Range("C7").NumberFormat = "+0.0 ""pbs"";-0.0 ""pbs"""
    End If

    ' YTM negociable
    ws.Range("A8").Value = "YTM negociable (%)"
    vAct = Val(wsH.Cells(ufH, 5).Value)
    ws.Range("B8").Value = vAct
    ws.Range("B8").NumberFormat = "0.0000"
    If tieneAnt Then
        vAnt = Val(wsH.Cells(ufH - 1, 5).Value)
        ws.Range("C8").Value = (vAct - vAnt) * 100
        ws.Range("C8").NumberFormat = "+0.0 ""pbs"";-0.0 ""pbs"""
    End If

    ' Spread
    ws.Range("A9").Value = "Spread negociable (pbs)"
    vAct = Val(wsH.Cells(ufH, 9).Value)
    ws.Range("B9").Value = vAct
    ws.Range("B9").NumberFormat = "#,##0.0"
    If tieneAnt Then
        vAnt = Val(wsH.Cells(ufH - 1, 9).Value)
        ws.Range("C9").Value = vAct - vAnt
        ws.Range("C9").NumberFormat = "+0.0;-0.0"
    End If

    ' % IG y rating
    ws.Range("A10").Value = "Investment grade / rating prom."
    ws.Range("B10").Value = Format(Val(wsH.Cells(ufH, 16).Value), "0.0%") & "  /  " & _
                            CStr(wsH.Cells(ufH, 17).Value)

    ' Concentracion
    ws.Range("A11").Value = "HHI / Top 5 / Top 10"
    ws.Range("B11").Value = Format(Val(wsH.Cells(ufH, 18).Value), "#,##0") & "  /  " & _
                            Format(Val(wsH.Cells(ufH, 19).Value), "0.0%") & "  /  " & _
                            Format(Val(wsH.Cells(ufH, 20).Value), "0.0%")

    ' Liquidez
    ws.Range("A12").Value = "Venc. 30d / 90d (MM)"
    ws.Range("B12").Value = Format(Val(wsH.Cells(ufH, 21).Value), "#,##0.0") & "  /  " & _
                            Format(Val(wsH.Cells(ufH, 22).Value), "#,##0.0")

    FormatoR ws.Range("A4:C12")

    ' ---------- Carry devengado ----------
    ws.Range("A14").Value = "CARRY DEVENGADO (no es el retorno del fondo)"
    ws.Range("A14").Font.Bold = True
    ws.Range("A15").Value = "MTD:"
    ws.Range("B15").Value = Val(wsH.Cells(ufH, 24).Value)
    ws.Range("B15").NumberFormat = "0.000%"
    ws.Range("A16").Value = "YTD:"
    ws.Range("B16").Value = Val(wsH.Cells(ufH, 25).Value)
    ws.Range("B16").NumberFormat = "0.000%"
    ws.Range("A14:B16").Font.Name = "Arial"
    ws.Range("A14:B16").Font.Size = 8

    ' ---------- Top 5 emisores ----------
    Dim wsC As Worksheet, i As Long
    Set wsC = HojaR(SH_CON)
    ws.Range("E4").Value = "TOP 5 EMISORES (negociable)": ws.Range("F4").Value = "PESO"
    For i = 1 To 5
        If Len(CStr(wsC.Cells(3 + i, 1).Value)) > 0 Then
            ws.Cells(4 + i, 5).Value = wsC.Cells(3 + i, 1).Value
            ws.Cells(4 + i, 6).Value = wsC.Cells(3 + i, 3).Value
            ws.Cells(4 + i, 6).NumberFormat = "0.00%"
        End If
    Next i
    FormatoR ws.Range("E4:F9")

    ' ---------- Calidad ----------
    Dim wsQ As Worksheet
    Set wsQ = HojaR(SH_CAL)
    ws.Range("H4").Value = "CALIDAD (negociable)": ws.Range("I4").Value = "PESO"
    For i = 1 To 6
        If Len(CStr(wsQ.Cells(3 + i, 1).Value)) > 0 Then
            ws.Cells(4 + i, 8).Value = wsQ.Cells(3 + i, 1).Value
            ws.Cells(4 + i, 9).Value = wsQ.Cells(3 + i, 3).Value
            ws.Cells(4 + i, 9).NumberFormat = "0.00%"
        End If
    Next i
    FormatoR ws.Range("H4:I10")

    ' ---------- Grafico 1: YTM + spread diario ----------
    GraficoTasa ws, wsH, ufH

    ' ---------- Grafico 2: vencimientos apilados ----------
    GraficoVencimientos ws

    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B:C").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 26
    ws.Columns("H").ColumnWidth = 12
    ws.Range("A1").Select

    Application.ScreenUpdating = True
    MsgBox "Resumen generado (" & Format(wsH.Cells(ufH, 1).Value, "dd/mm/yyyy") & ").", vbInformation
End Sub

'------------------------------------------------------------
Private Sub GraficoTasa(ws As Worksheet, wsH As Worksheet, ByVal ufH As Long)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=ws.Range("A18").Left, Top:=ws.Range("A18").Top, _
                                 Width:=540, Height:=230)
    With co.Chart
        .ChartType = xlLine
        Dim s As Series
        Set s = .SeriesCollection.NewSeries
        s.Name = "YTM fondo (%)"
        s.XValues = wsH.Range("A2:A" & ufH)
        s.Values = wsH.Range("D2:D" & ufH)
        s.Format.Line.ForeColor.RGB = RGB(212, 12, 12)
        s.Format.Line.Weight = 1.75

        Set s = .SeriesCollection.NewSeries
        s.Name = "Spread (pbs)"
        s.XValues = wsH.Range("A2:A" & ufH)
        s.Values = wsH.Range("I2:I" & ufH)
        s.AxisGroup = xlSecondary
        s.Format.Line.ForeColor.RGB = RGB(191, 168, 128)
        s.Format.Line.Weight = 1.25

        .HasTitle = True
        .ChartTitle.Text = "Tasa ponderada diaria — YTM (izq) vs Spread (der)"
        .ChartTitle.Font.Size = 9
        .ChartTitle.Font.Name = "Arial"
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 7
        .Axes(xlValue).TickLabels.Font.Size = 7
        .Axes(xlValue, xlSecondary).TickLabels.Font.Size = 7
        .Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"
    End With
End Sub

'------------------------------------------------------------
Private Sub GraficoVencimientos(ws As Worksheet)
    Dim wsV As Worksheet, ufV As Long
    Set wsV = HojaR(SH_VEN)
    ufV = UltFilaR(wsV)
    If ufV < 4 Then Exit Sub

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=ws.Range("A18").Left + 560, Top:=ws.Range("A18").Top, _
                                 Width:=380, Height:=230)
    With co.Chart
        .ChartType = xlColumnStacked
        Dim s As Series

        Set s = .SeriesCollection.NewSeries
        s.Name = "Bonos"
        s.XValues = wsV.Range("A4:A" & ufV)
        s.Values = wsV.Range("B4:B" & ufV)
        s.Format.Fill.ForeColor.RGB = RGB(212, 12, 12)

        Set s = .SeriesCollection.NewSeries
        s.Name = "Dep. Plazo"
        s.XValues = wsV.Range("A4:A" & ufV)
        s.Values = wsV.Range("C4:C" & ufV)
        s.Format.Fill.ForeColor.RGB = RGB(89, 89, 89)

        Set s = .SeriesCollection.NewSeries
        s.Name = "Papeles Com."
        s.XValues = wsV.Range("A4:A" & ufV)
        s.Values = wsV.Range("D4:D" & ufV)
        s.Format.Fill.ForeColor.RGB = RGB(191, 168, 128)

        .HasTitle = True
        .ChartTitle.Text = "Vencimientos por año (S/ MM)"
        .ChartTitle.Font.Size = 9
        .ChartTitle.Font.Name = "Arial"
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 7
        .Axes(xlValue).TickLabels.Font.Size = 7
    End With
End Sub

'------------------------------------------------------------
Private Sub BorrarGraficos(ws As Worksheet)
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co
End Sub

Private Sub FormatoR(rg As Range)
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

Private Function HojaR(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    End If
    Set HojaR = ws
End Function

Private Function UltFilaR(ws As Worksheet) As Long
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then UltFilaR = 0 Else UltFilaR = c.Row
End Function
