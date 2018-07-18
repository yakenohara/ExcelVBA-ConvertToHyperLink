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
    
    Dim cautionMessage As String: cautionMessage = "����Sub�v���V�[�W���́A" & vbLf & _
                                                   "���݂̑I��͈͂ɑ΂��Ēl�̏������݂��s���܂��B" & vbLf & vbLf & _
                                                   "���s���܂���?"
    
    '���s�m�F
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
    End If
    
    
    '�V�[�g�I����ԃ`�F�b�N
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    '������
    numOfCells = Selection.count
    
    '���s���[�v
    cellcnt = 1
    For Each writePlace In Selection
        
        Application.StatusBar = "processing " & cellcnt & " of " & numOfCells
        
        If (writePlace.Address = writePlace.MergeArea.Cells(1, 1).Address) Then '�����Z���łȂ��ꍇ
            
            val = writePlace.MergeArea.Cells(1, 1).Value
            Set writePlace = writePlace.MergeArea
            
            If (Not (val = "")) Then 'vacant�łȂ��ꍇ
            
            
                '�n�C�p�[�����N�̍쐬
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



