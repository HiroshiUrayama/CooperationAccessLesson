Attribute VB_Name = "KeyWordNew"
'#################################################
'New�L�[���[�h�̂��׋�
'#################################################
Option Explicit

'==================================================
'������New����Ɖ������Ȃ̂�(�Ƃ肠�����ǂ��������������)
'==================================================
Private Sub ADOSample4()
    
#Const SAME = True

#If SAME Then
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
#Else
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
#End If
    
    
    'True�̏ꍇ�A�����̎��_�ł͂܂��I�u�W�F�N�g���̂���������Ă��Ȃ�
    'FALSE�̏ꍇ�́A�����̎��_�ł��łɃI�u�W�F�N�g����������Ă���
    Stop
    
    '�������͈̂ꏏ�Ȃ̂ŁA�ϐ��̐錾�����������߂���
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\�ڋq�f�[�^.accdb"
                
    rs.Open "T�ڋq���X�g", cn, adOpenForwardOnly, adLockReadOnly
    
    MsgBox "�y1���ڂ̃f�[�^�z" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf
        
    Stop
    
    rs.Close
    cn.Close
End Sub

'==================================================
'������New����Ƃ܂����p�^�[��
'==================================================
Private Sub NewKeywordSample1()
    
#Const ARI = True

#If ARI Then
    '�y���[�J���E�C���h�E�Ŋm�F����Έꔭ�z
    '���̏������́A�ϐ��̐錾���ɃI�u�W�F�N�g�𐶐�����킯�ł͂Ȃ��B
    '�g�p����ۂɁu�Ȃ���΍��v�Ƃ�������������(����ȍ~�������悤�ȓ���������)
    '�ϐ����g�p���邽�тɕϐ���]�����鏈�������邩��A�킸�������ǒx���Ȃ�B
    Dim obj As New Collection
#Else
    Dim obj As Object
    Set obj = New Collection
#End If
    
    '�v�f��ǉ�����
    obj.Add "A"
    
    'Collection�I�u�W�F�N�g�ւ̎Q�Ƃ���������
    '�yARI��True�̂Ƃ��z���ꂪ�ԈႢ�Ȃ̂��A���̌��Add���ԈႢ�Ȃ��̂����킩��Ȃ��Ȃ�
    Set obj = Nothing
    
    '�ēx�v�f��ǉ�
    'ARI=True�̂Ƃ��͏����obj�����o�����B��add�����(Nothing����Ă���̂ɍēx��������Ă��܂�)
    'ARI=FALSE�̂Ƃ��͏����obj�����o���ꂸ�A�����ł�����ƃG���[��f���Ď~�܂�(Nothing���ꂽ���ʂ���)
    obj.Add "B"
    Stop
End Sub

'������������������������������������������������������������
'
'������������������������������������������������������������
