Attribute VB_Name = "ConvertToHyperLink"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

Sub ConvertToHyperLink()
    
    Dim writePlace As Range
    Dim val As Variant
    Dim retVal As Integer
    
    Dim numOfCells As Long
    Dim cellcnt As Long
    
    Dim cautionMessage As String: cautionMessage = "このSubプロシージャは、" & vbLf & _
                                                   "現在の選択範囲に対して値の書き込みを行います。" & vbLf & vbLf & _
                                                   "実行しますか?"
    
    '実行確認
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
    End If
    
    
    'シート選択状態チェック
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "複数シートが選択されています" & vbLf & _
               "不要なシート選択を解除してください"
        Exit Sub
    End If
    
    '初期化
    numOfCells = Selection.count
    
    '実行ループ
    cellcnt = 1
    For Each writePlace In Selection
        
        Application.StatusBar = "processing " & cellcnt & " of " & numOfCells
        
        If (writePlace.Address = writePlace.MergeArea.Cells(1, 1).Address) Then '結合セルでない場合
            
            val = writePlace.MergeArea.Cells(1, 1).Value
            Set writePlace = writePlace.MergeArea
            
            If (Not (val = "")) Then 'vacantでない場合
            
            
                'ハイパーリンクの作成
                ActiveSheet.Hyperlinks.Add _
                                        Anchor:=writePlace, _
                                        Address:=val, _
                                        TextToDisplay:="'" & val
            
            End If
            
        End If
        
        cellcnt = cellcnt + 1
        
    Next writePlace
    
    Application.StatusBar = False
    
    MsgBox "Done!"
    
End Sub



