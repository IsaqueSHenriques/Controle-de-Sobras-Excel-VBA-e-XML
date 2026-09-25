Attribute VB_Name = "Módulo11"
Sub lerXML()
    Dim dialogo As FileDialog
    Dim arquivo As Variant
    Dim xmlDoc As Object
    Dim AddedPart As Object
    Dim Part As Object
    Dim Piece As Object
    Dim Sid As Object
    Dim peca As CAddedPart
    Dim atributos As Object
    Dim pecas As Collection
    Dim Aid As Object
    Dim AidQ As String
    Dim atributosP As Object
    Dim i As Long
    Dim i2 As Long
    Dim Partm As Object
    Dim material As String
    Dim CutList As Object
    Dim lote As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ultimaColuna As Long
    Dim nameColumn() As Variant
    Dim dataCorte As String
     
    Set dialogo = Application.FileDialog(msoFileDialogFilePicker)
    
    dialogo.AllowMultiSelect = True
    dialogo.Filters.Add "Arquivo XML", "*.xml"
    
    dataCorte = InputBox("Insira a data de corte: ")
    
    If (dialogo.Show = -1) Then
    
        Set wb = Workbooks("1- Sobras Guarda - Roupa.xlsx")
        
            For Each arquivo In dialogo.SelectedItems
            
                Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
                xmlDoc.Load arquivo
                
                Set CutList = xmlDoc.GetElementsByTagName("CutList")
                lote = CutList.Item(0).GetAttribute("order")
                lote = Replace(lote, "LOTE ", "")
                lote = Replace(lote, ".opj", "")
                
                Set Partm = xmlDoc.GetElementsByTagName("Part")
                material = Partm.Item(0).GetAttribute("Material")
                
                Set ws = wb.Worksheets(material)
                
                Set AddedPart = xmlDoc.GetElementsByTagName("AddedPart")
                
                Set pecas = New Collection
                
                    For Each atributos In AddedPart
                    
                        Set peca = New CAddedPart
                        peca.IP = atributos.GetAttribute("id")
                        peca.Desc = atributos.GetAttribute("IDesc")
                        peca.Comp = CLng(Val(atributos.GetAttribute("L")))
                        peca.Larg = CLng(Val(atributos.GetAttribute("W")))
                        pecas.Add peca
                        
                    Next atributos

                    Set Piece = xmlDoc.GetElementsByTagName("Piece")

                        For Each atributosP In Piece

                            If atributosP.GetElementsByTagName("Aid").Length > 0 Then

                                Set Aid = atributosP.GetElementsByTagName("Aid")(0)
                                AidQ = Aid.GetAttribute("id")

                                    For i = 1 To pecas.Count

                                        If pecas(i).IP = AidQ Then

                                            pecas(i).QTDD = CLng(atributosP.GetAttribute("Q"))

                                    Exit For

                                        End If

                                    Next i

                                End If

                            Next atributosP
                    
                    For i = 1 To pecas.Count
            
                        If pecas(i).Desc = "SOBRA" Then
                
                            ultimaLinha = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row + 1
                            
                            Do While ws.Cells(ultimaLinha, 6).Value = ""

                                ultimaLinha = ultimaLinha - 1

                            Loop
                            
                            ultimaLinha = ultimaLinha + 1
                            
                            If ultimaLinha <> 5 Then
                                
                                ws.Rows(ultimaLinha - 1).Copy
                                ws.Rows(ultimaLinha).PasteSpecial _
                                Paste:=xlPasteFormats
                                Application.CutCopyMode = False
                            
                            End If
                            
                            
                            ws.Rows(ultimaLinha).RowHeight = _
                            ws.Rows(ultimaLinha - 1).RowHeight
                            
                            ultimaColuna = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
                
                            ReDim nameColumn(3 To ultimaColuna)
                
                                For i2 = 3 To ultimaColuna
                        
                                    nameColumn(i2) = ws.Cells(4, i2).Value
                                    
                                        If nameColumn(i2) = "LOTE" Then
                                        
                                            ws.Cells(ultimaLinha, i2).Value = lote
                                            ws.Cells(ultimaLinha, i2).Interior.Color = RGB(169, 208, 142)
                                            
                                                ElseIf nameColumn(i2) = "COMP." Then
                                                
                                                    ws.Cells(ultimaLinha, i2).Value = pecas(i).Comp
                                                                                                                              
                                                        ElseIf nameColumn(i2) = "LARG." Then
                                                        
                                                            ws.Cells(ultimaLinha, i2).Value = pecas(i).Larg
                                                            
                                                                ElseIf nameColumn(i2) = "QUANTIDADE" Then
                                                                
                                                                    ws.Cells(ultimaLinha, i2).Value = pecas(i).QTDD
                                                                    
                                                                        ElseIf nameColumn(i2) = "DESCRIÇÃO" Then
                                                                        
                                                                            ws.Cells(ultimaLinha, i2).Value = pecas(i).Desc
                                                                            
                                                                                ElseIf nameColumn(i2) = "DATA DE CORTE" Then
                                                                                
                                                                                    ws.Cells(ultimaLinha, i2).Value = dataCorte
                                                                            
                                        End If
                                                                
                                Next i2
                                                                                   
                        End If
                        
                    Next i
                       
                Next arquivo
                
                    
                MsgBox ("IMPORTAÇÃO COMPLETA.")
                
    End If
    

    
End Sub



