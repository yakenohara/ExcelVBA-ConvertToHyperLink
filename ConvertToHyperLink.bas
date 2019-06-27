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
    
    Dim numOfCells As LongLong
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
    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "複数シートが選択されています" & vbLf & _
               "不要なシート選択を解除してください"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '初期化
    numOfCells = Selection.CountLarge
    
    'シート範囲全選択されていた場合は、UsedRange内に収まるようにトリミング
    Set range_selection = trimWithUsedRange(Selection)
    
    '実行ループ
    cellcnt = 1
    For Each writePlace In range_selection
        
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
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "Done!"
    
End Sub

'
' セル参照範囲が UsedRange 範囲に収まるようにトリミングする
'
Private Function trimWithUsedRange(ByVal rangeObj As Range) As Range

    'variables
    Dim ret As Range
    Dim long_bottom_right_row_idx_of_specified As Long
    Dim long_bottom_right_col_idx_of_specified As Long
    Dim long_bottom_right_row_idx_of_used As Long
    Dim long_bottom_right_col_idx_of_used As Long

    '指定範囲の右下位置の取得
    long_bottom_right_row_idx_of_specified = rangeObj.Item(1).Row + rangeObj.Rows.Count - 1
    long_bottom_right_col_idx_of_specified = rangeObj.Item(1).Column + rangeObj.Columns.Count - 1
    
    'UsedRangeの右下位置の取得
    With rangeObj.Parent.UsedRange
        long_bottom_right_row_idx_of_used = .Item(1).Row + .Rows.Count - 1
        long_bottom_right_col_idx_of_used = .Item(1).Column + .Columns.Count - 1
    End With
    
    'トリミング
    Set ret = rangeObj.Parent.Range( _
        rangeObj.Item(1), _
        rangeObj.Parent.Cells( _
            IIf(long_bottom_right_row_idx_of_specified > long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_used, long_bottom_right_row_idx_of_specified), _
            IIf(long_bottom_right_col_idx_of_specified > long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_used, long_bottom_right_col_idx_of_specified) _
        ) _
    )
    
    '格納して終了
    Set trimWithUsedRange = ret
    
End Function

