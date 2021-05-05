Attribute VB_Name = "DataControl_important"
'#################################################
'���x�ȏ���(�f�[�^�x�[�X��)
'#################################################
Option Explicit

'�d�v
'�f---------------------------------------------------
'�f�[�^�̐�������ۂ���
'�~�X�Ńf�[�^���Ԉ�����l�ɍX�V����Ȃ��悤�ɂ��邱��

'OpenRecordset���\�b�h(DAO)�̈���LockEdit�Ɏw�肷��LockTypeEnum�񋓌^�̒l
'==================================================
'dbOptimistic(3)
'���R�[�hID�ɂ��Ǔ˂����L�I�������b�N�B�J�[�\���͌Â����R�[�h�ƐV�������R�[�h�̃��R�[�hID���r���A
'���̃��R�[�h�ւ̃A�N�Z�X���Ō�ɍs���Ă���ύX��������ꂽ���ǂ������f����

'dbPessimistic(2)
'�r���I�������b�N�B
'�J�[�\���́A���R�[�h���X�V�\�ł��邱�Ƃ�ۏ؂��邽�߂ɕK�v�ȍŒ���̃��b�N���g�p����B

'Open���\�b�h(ADO)�̈���LockType�Ɏw�肷��LockTypeEnum�񋓌^�̒l
'==================================================
'adLcokReadOnly(1)
'����l�B�ǂݎ���p�B�f�[�^�̕ύX�s�B

'adLockPessimistic(2)
'���R�[�h�P�ʂ̔r���I���b�N�B�ҏW����̃f�[�^�\�[�X�Ń��R�[�h�����b�N����B

'adLockOptimistic(3)
'���R�[�h�P�ʂ̋��L�I���b�N�B
'Update���\�b�h���Ăяo�����ꍇ�ɂ̂݁A���R�[�h�����b�N����B

'adLockBatchOptimistic(4)
'���L�I�o�b�`�X�V�B�o�b�`�X�V���[�h�̏ꍇ�ɂ̂ݎw��\�B

'adLockUnspecified(-1)
'���b�N�̎�ނ��w�肵�Ȃ��B

'���R�[�h�P�ʂ̃��b�N���\
'(ADO�̏ꍇ��)�ǂݎ���p�ŊJ���邱��
'���̎d�g�݂𗝉����Ă������Ƃ��d�v

'==================================================
'�yDAO�z�g�����U�N�V��������
'�r���ŃG���[�ɂȂ����ꍇ�̕s�������Ȃ�������(DAO�ɂ�ADO�ɂ�����)
'==================================================
Private Sub DAOTransactionSample()
#Const ARI = True

#If ARI Then
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
#Else
    Dim ws As Object
    Dim db As Object
    Dim rs As Object
#End If

    Set ws = DAO.DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase("C:\Users\USER\Desktop\Git�Ǘ���\CooperationAccessLesson" & "\����f�[�^.accdb")
    
    '�e�[�u���uT_���i�}�X�^�v�����R�[�h�Z�b�g���ĊJ��
    Set rs = db.OpenRecordset("T_���i�}�X�^", dbOpenDynaset)

    '�G���[�������J�n����
    On Error GoTo ErrHdl
    
    '�g�����U�N�V�����������J�n---------------------------------------------------
    ws.BeginTrans
    
    rs.AddNew
    'rs!ID = "B0004"
    rs!���i�� = "�x���gD"
    rs!�P�� = 12000
    rs.Update
    
    ws.CommitTrans
    '�������m��---------------------------------------------------
    
ExitHdl:
    rs.Close
    db.Close
    ws.Close
    Exit Sub

ErrHdl:
    ws.Rollback
End Sub

'==================================================
'�yADO�z�g�����U�N�V��������
'�r���ŃG���[�ɂȂ����ꍇ�̕s�������Ȃ�������(DAO�ɂ�ADO�ɂ�����)
'==================================================
Private Sub ADOTransactionSample()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'Connectiion�I�u�W�F�N�g���쐬
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                 "Data Source=" & "C:\Users\USER\Desktop\Git�Ǘ���\CooperationAccessLesson" & "\����f�[�^.accdb;"
    
    rs.Open "T_���i�}�X�^", cn, adOpenDynamic, adLockOptimistic
    
    '�G���[�����J�n
    On Error GoTo ErrHdl
    
    cn.BeginTrans '---------------------------------------------------
    rs.AddNew
    'rs!ID =""
    rs!���i�� = "�p���cA"
    rs!�P�� = 15000
    rs.Update
    cn.CommitTrans '---------------------------------------------------
    
ExitHdl:
    rs.Close
    cn.Close
    Exit Sub
    
ErrHdl:
    cn.RollbackTrans
End Sub

'���̂�����̃G���[�����A�g�����U�N�V�����̓Z�b�g�Ŏg�p����(�g����悤�ɂȂ��Ă����ق����x�X�g)

'---------------------------------------------------
'Access���̂��̂̏���������������ɂ�
'�Q�Ɛݒ�yMicrosoft Access XX.X Object Library�z
'DoCmd�I�u�W�F�N�g���g�p����
'Access�t�H�[�����J���A����A�N�G�������s����Ƃ���������͂ł��Ă��A
'�f�[�^�𒼐ڍX�V�����肷�邱�Ƃ͂ł��Ȃ�
'---------------------------------------------------














































