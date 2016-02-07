Imports FPSpread.CellTypeConstants
Imports AxFPSpread

Module modSpread
    Public Sub LimpiaBloque(ByRef ctlSpread As AxFPSpread.AxvaSpread, ByVal vintRen%, ByVal vintCol%, ByVal vintRen2%, ByVal vintCol2%)
        '*************************************************************************
        ' Rutina que limpia los DATOS de de las celdas tipo TEXTO un spread
        ' en el rango dado.
        ' Entrada :
        '   ctlSpread .- Spread
        '   vintRen .- Renglón inicio
        '   vintCol .- Columna inicio
        '   vintRen2 .- Renglon final
        '   vintCol2 .- Columna final
        '*************************************************************************
        Dim i As Integer

        ctlSpread.ClearRange(vintCol, vintRen, vintCol2, vintRen2, True)
        ctlSpread.SetActiveCell(vintCol, vintRen)

        For i = vintRen To vintRen2
            ctlSpread.Col = 1
            ctlSpread.Row = i

            If ctlSpread.CellType = CellTypePicture Then
                ctlSpread.DeleteRows(i, 1)
            End If

        Next i
    End Sub

    Sub MakeFloatCell(ByVal Col As Long, ByVal col2 As Long, ByVal Row As Long, ByVal row2 As Long, ByVal floatmin As String, _
    ByVal floatmax As String, ByVal floatmoney As Boolean, ByVal floatsep As Boolean, ByVal decplaces As Integer, ByVal fpvalue As Double, _
    ByRef ctlSpread As AxFPSpread.AxvaSpread)

        ctlSpread.Col = Col
        ctlSpread.Col2 = col2
        ctlSpread.Row = Row
        ctlSpread.Row2 = row2
        ctlSpread.BlockMode = True
        'Define cells as type FLOAT
        If floatmoney Then
            ctlSpread.CellType = CellTypeCurrency
            ctlSpread.TypeCurrencyShowSymbol = True
            ctlSpread.TypeCurrencyDecPlaces = decplaces
            ctlSpread.TypeCurrencyShowSep = floatsep
            ctlSpread.TypeCurrencyMin = floatmin
            ctlSpread.TypeCurrencyMax = floatmax
        Else
            ctlSpread.CellType = CellTypeNumber
            ctlSpread.TypeNumberDecPlaces = decplaces
            ctlSpread.TypeNumberShowSep = floatsep
            ctlSpread.TypeNumberMin = floatmin
            ctlSpread.TypeNumberMax = floatmax
        End If
        ctlSpread.Value = fpvalue
        ctlSpread.BlockMode = False

    End Sub
End Module
