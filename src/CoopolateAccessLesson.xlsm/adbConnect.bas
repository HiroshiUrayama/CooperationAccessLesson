Attribute VB_Name = "adbConnect"
'#################################################
'ADO��Access�ɐڑ�����
'#################################################
Option Explicit

'==================================================
'ADO�𗘗p���ăf�[�^�x�[�X�ɐڑ�����R�[�h
'�܂����́AModule���̂�ADO�Ƃ������Ă���ƎQ�Ɛݒ肪�ł���(���łɎg���Ă��܂��\��)
'==================================================
Private Sub dbSample()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    'Connection�I�u�W�F�N�g���쐬����
    Set cn = New ADODB.Connection
    
    'Recordset�I�u�W�F�N�g���쐬����
    Set rs = New ADODB.Recordset
    
    '�u�ڋq�f�[�^.accdb�v�f�[�^�x�[�X���J��
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\�ڋq�f�[�^.accdb;"
                
    '�uT�ڋq���X�g�v�e�[�u���̃f�[�^���擾����(adOpenForwardOnly,adLockReaOnly������l)
    rs.Open "T�ڋq���X�g", cn, adOpenForwardOnly, adLockReadOnly
    
    '���b�Z�[�W���擾����
    MsgBox "�y1���ڂ̃f�[�^�z" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(4).Name & ":" & rs.Fields(1).Value & vbCrLf
        
    rs.Close
    cn.Close

End Sub

'---------------------------------------------------
'���R�[�h�Z�b�g�̊J��������؁B
'---------------------------------------------------
'object.Open Source, ActiveConnection, CursorType, LockType, Option
'�I�u�W�F�N�g��cn(Connection�I�u�W�F�N�g)���w�肷��ƁA�f�[�^�x�[�X�ɐڑ����邱�Ƃ��ł���B

'Source�F�ȗ��\�B�e�[�u������N�G�����ASQL�����w�肷��B
'ActiveConnection�F�ȗ��\�BConnection�I�u�W�F�N�g�܂��͐ڑ���񕶎�����w�肷��B
'CursorType�F�ȗ��\�B�J�[�\���^�C�v�����߂邽�߂̒l���w�肷��B
'LockType�F�ȗ��\�B���b�N�̎�ނ����߂邽�߂̒l���w�肷��B
'Options�F�ȗ��\�B

'---------------------------------------------------
'�J��������؁B
'---------------------------------------------------
'��CursorType
'adOpenForwardOnly(0)�^(�O����p�J�[�\��)
    '����l�B���R�[�h�̃X�N���[���������O�����Ɍ��肳��Ă��邱�Ƃ������A�ÓI�J�[�\���Ɠ�������������
'adOpenKeyset(1)�^(�L�[�Z�b�g�J�[�\��)
    '���̃��[�U���ǉ��������R�[�h�͕\���ł��Ȃ��B����ȊO�͓��I�J�[�\���Ɠ����B
'adOpenDynamic(2)�^(���I�J�[�\��)
    '���̃��[�U�ɂ��ǉ��A�ύX�A�y�э폜���m�F�ł���B
    '�v���o�C�_���u�b�N�}�[�N���T�|�[�g���Ă���ꍇ�ARecordset���ł̑S�Ă̑����������
'adOpenStatic(3)�^�f�[�^�̌����⃌�|�[�g�̍쐬�Ɏg�p���邽�߂́A���R�[�h�̐ÓI�R�s�[�B
    '���̃��[�U�[�ɂ��ǉ��A�ύX�A�폜�͕\������Ȃ��B
'adOpenUnspecified(-1)
    '�J�[�\���̎�ނ��w�肵�Ȃ�
    
'��LockType
'adLockReadOnly(1)�^����l�B�ǂݎ���p�B�f�[�^�̕ύX�s��
'adLockPessimistic(2)�^���R�[�h�P�ʂ̔r���I���b�N�B
                                '�ҏW����̃f�[�^�\�[�X�Ń��R�[�h�����b�N����B
'adLockOptimistic(3)�^���R�[�h�P�ʂ̋��L�I���b�N�BUpdate���\�b�h���Ăяo�����ꍇ�ɂ̂݃��R�[�h�����b�N�B
'adLockBatchOptimistic(4)�^���L�I�o�b�`�X�V�B
                                '�o�b�`�X�V���[�h�̏ꍇ�ɂ̂ݎw��\�B
'adLockUnspecified(-1)�^���b�N�̎�ނ��w�肵�Ȃ�

