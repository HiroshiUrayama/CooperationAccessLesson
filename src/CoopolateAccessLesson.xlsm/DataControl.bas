Attribute VB_Name = "DataControl"
'#################################################
'SQL
'#################################################
Option Explicit

'==================================================
'�N�G�������s����R�[�h
'MicrosoftDAO3.6�I�u�W�F�N�g���L���ɂȂ��Ă����Microsoft 16.0 Data(DAO)�Ɨ����ł��Ȃ�(�ǂ�����Q�Ƃ���̂�"DAO�Ȃ̂�")
'==================================================
Private Sub DAOQuerySample()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Long
    
    Set ws = DAO.DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(ThisWorkbook.Path & "\�ڋq�f�[�^.accdb")
    
    '�N�G��������\���Ă��āA�ǂ�ȏ������s���̂����킩��Ȃ��B
    'Access���J���ɍs���āA�N�G���r���[�����Ȃ��Ɓc
    Set rs = db.OpenRecordset("Q�����s")
    
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

'==================================================
'SQL���g�p������
'==================================================
Private Sub SQLSample1()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim vSQL As String
    Dim i As Long
    
    '���[�N�X�y�[�X���`
    Set ws = DAO.DBEngine.Workspaces(0)
    
    '�u�ڋq�f�[�^.accdb�ɐڑ�����v
    Set db = ws.OpenDatabase("C:\Users\USER\Desktop\Git�Ǘ���\CooperationAccessLesson" & "\�ڋq�f�[�^.accdb")
    
    '�uQ�����s�v�e�[�u���̃��R�[�h�Z�b�g���J��
    '���o����SQL����������
    'SQL���ƃG�f�B�^�����牽����낤�Ƃ��Ă���̂����ꔭ�ŕ�����
    vSQL = "SELECT * FROM T�ڋq���X�g WHERE �s���{�� = '�����s'"
    
    '�w�肵��sql�����g���ă��R�[�h�Z�b�g���J��
    Set rs = db.OpenRecordset(vSQL, dbOpenDynaset)
    
    '���R�[�h�Z�b�g�̑S�Ẵt�B�[���h�ɑ΂��ď������s��
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).Name
    Next
    
    rs.Close
    db.Close
    ws.Close
End Sub

'==================================================
'Access����f�[�^���擾���ă��[�N�V�[�g�ɓ\�t����T���v��
'==================================================
Private Sub SQLSample2()
    
#Const ARI = True

#If ARI Then
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
#Else
    Dim cn As Object
    Dim rs As Object
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
#End If

    '�R�l�N�V�����I�[�v��
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                 "Data Source=" & "C:\Users\USER\Desktop\Git�Ǘ���\CooperationAccessLesson" & "\�ڋq�f�[�^.accdb"

    '35�Έȏ�̃f�[�^���擾����SQL����ݒ肷��
    With rs
        .ActiveConnection = cn
        .Source = "SELECT * FROM T�ڋq���X�g WHERE �N�� > 35"
        .Open
    End With
    
    '�t�B�[���h������͂���
    Dim i As Long
    For i = 0 To rs.Fields.Count - 1
        Cells(1, i + 1).Value = rs.Fields(i).Name
    Next
    
    '�擾�����f�[�^���Z��A2�ȍ~�ɓ\�t����
    '����A�֗���ȁc
    Range("A2").CopyFromRecordset rs
        
    '�R�l�N�V��������
    cn.Close
    
    '�񕝎�������
    Columns("A:K").AutoFit
    
End Sub

'==================================================
'Access����f�[�^���擾���ă��[�N�V�[�g�ɓ\�t����T���v��(Update��)
'==================================================
Private Sub UpdateSample()
#Const ado = True

#If ado Then
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
#Else
    Dim cn As Object
    Set cn = CreateObject("ADODB.Conneciton")
#End If
    
    Dim vSQL As String
    
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & "C:\Users\USER\Desktop\Git�Ǘ���\CooperationAccessLesson" & "\�ڋq�f�[�^.accdb"
    
    vSQL = "UPDATE T�ڋq���X�g SET �N�� = 48 WHERE �ڋq�� = '�c�� �m�s'"

    '�ݒ肵��SQL�������s����
    cn.Execute vSQL
    
    cn.Close
    
End Sub

'#If ARI Then
'#Else
'#End If






























































