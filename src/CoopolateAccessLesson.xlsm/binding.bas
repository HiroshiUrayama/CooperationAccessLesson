Attribute VB_Name = "binding"
'#################################################
'���O�o�C���f�B���O�Ǝ��s���o�C���f�B���O�̈Ⴂ(Code)
'#################################################
Option Explicit

'==================================================
'���O�o�C���f�B���O
'==================================================
Private Sub sample1()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    '�u�ڋq�f�[�^.accdb�v�f�[�^�x�[�X���J��
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\�ڋq�f�[�^.accdb;"
    
    '���R�[�h�Z�b�g�����o��
    rs.Open "T�ڋq���X�g", cn, adOpenForwardOnly, adLockReadOnly
    
    '���b�Z�[�W�\��
    MsgBox "�y1���ڂ̃f�[�^�z" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf

    rs.Close
    cn.Close
End Sub

'==================================================
'���s���o�C���f�B���O
'==================================================
Private Sub sample2()
    
    '�ϐ��錾��Object�^�ɂȂ��Ă���
    Dim cn As Object
    Dim rs As Object
    
    'CreateObject�֐����g�p
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    '�u�ڋq�f�[�^.accdb�v�f�[�^�x�[�X���J��
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\�ڋq�f�[�^.accdb;"
    
    '���R�[�h�Z�b�g�����o��
    rs.Open "T�ڋq���X�g", cn, adOpenForwardOnly, adLockReadOnly
    
    '���b�Z�[�W�\��
    MsgBox "�y1���ڂ̃f�[�^�z" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf

    rs.Close
    cn.Close
End Sub

'������������������������������������������������������������
'���O�o�C���f�B���O�Ǝ��s���o�C���f�B���O�̃����b�g�A�f�����b�g
'������������������������������������������������������������
'�o�C���f�B���O             �f�����b�g                              �f�f�����b�g
'���O�o�C���f�B���O         �f�������x���኱����            �f�ėp���ɂ�����P�[�X���L��
                                '�C���e���Z���X�����p�ł���
'���s���o�C���f�B���O       �f�������x����r�I�x��          �f�ėp��������
                                '�C���e���Z���X���g���Ȃ�

'���܂�A�J������Ƃ��͎��O�o�C���f�B���O�Ǝ��s���o�C���f�B���O�̂����Ƃ���������΂����킯�ŁB
'������ǂ�����č��̂����l����̂���؁B

'---------------------------------------------------
'�ėp���Ƃ́F�쐬�����v���O�������𗘗p����ۂɁA�ʂ�PC�ł�������Ɠ��삷�邩�H
'---------------------------------------------------
'�Г��Ŏg�p����ۂɁA�l�ɂ���Ďg�p����Office�̃o�[�W������������肷��ƁA�Q�Ɛݒ肪�g���Ȃ��Ȃ�
'�Q�ƕs�ɂȂ�����A�Q�Ɛݒ�𒲐߂����炢�����ǁAVBA���g���Ă��Ȃ��l�͎Q�Ɛݒ��m��Ȃ��l�������̂ŁA�蓮�ł��̂����
'CreateObject�֐��ł��ƍ��̃o�[�W�����ɍ��킹�Ă����̂ŁA���ꂪ�u�ėp���������v�Ƃ����Ӗ��B

'VBA�ŎQ�Ɛݒ�̗L���𒲂ׂ��邪�A�u�Z�L�����e�B�Z���^�[�v�ŁuVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v�Ƀ`�F�b�N������K�v������
'�Q�Ɛݒ���������郁�\�b�h�͂��邪(Remove)�A�Q�ƕs�̃��C�u�����ɑ΂���Remove���\�b�h�����s����ƁA�G���[�ɂȂ�
'���Q�ƕs�̏ꍇ�͎蓮�ŉ������邵�������B

'==================================================
'�y�ŏI�`�z
'���O�o�C���f�B���O�Ǝ��s���o�C���f�B���O�̂����Ƃ�����������Ƃ̎d��
'�J������Ƃ��͎��O�o�C���f�B���O�ŊJ��
'���삳����Ƃ��͎��s���o�C���f�B���O�œ��삳����
'���J�����ɂ̓C���e���Z���X���g���Ă������ƊJ�����āA�����[�X���ɂ͎��s���o�C���f�B���O�Ŕėp�����������ĊJ������
    '�����[�X���ɂ͎Q�Ɛݒ���O���葱���͕K�v�B
'==================================================
Private Sub ADOSample3()
    
'�����t���R���p�C���萔(���s���ŕύX�����肷��)
#Const HAS_REF = True

'�����t���R���p�C���̋L���������Ă����Ȃ��ƁA�R���p�C����S�s�ɘj���čs���̂ŃG���[�ɂȂ�͂�
#If HAS_REF Then
    '�Q�Ɛݒ肪ON�̂Ƃ��͂����炪���������(�܂�A�J�����̎�)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
#Else
    '�Q�Ɛݒ肪OFF�̂Ƃ��͂����炪���������(�܂�A�����[�X�̎�)
    Dim cn As Object
    Dim rs As Object
    cn = CreateObject("ADODB.Connection")
    rs = CreateObject("ADODB.Recordset")
#End If
    
    '�������͈̂ꏏ�Ȃ̂ŁA�ϐ��̐錾�����������߂���
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                "Data Source=" & ThisWorkbook.Path & "\�ڋq�f�[�^.accdb"
                
    rs.Open "T�ڋq���X�g", cn, adOpenForwardOnly, adLockReadOnly
    
    MsgBox "�y1���ڂ̃f�[�^�z" & vbCrLf & _
        rs.Fields(0).Name & ":" & rs.Fields(0).Value & vbCrLf & _
        rs.Fields(1).Name & ":" & rs.Fields(1).Value & vbCrLf & _
        rs.Fields(2).Name & ":" & rs.Fields(2).Value & vbCrLf & _
        rs.Fields(3).Name & ":" & rs.Fields(3).Value & vbCrLf

End Sub

'---------------------------------------------------
'�f�B���N�e�B�u(�����t���R���p�C��)
'---------------------------------------------------
'�u#Const �f�B���N�e�B�u�v�Œ萔��錾���A�u#If...Then...#Else �f�B���N�e�B�u�v�ŏ����𕪂���
'#�����Ă镔���͕���������R���p�C�����Ȃ�(#�萔��True�ł���΁A#If����True�̕����݂̂����R���p�C�����Ȃ�)
'��True�AFalse�̎��ɕϐ��̖��O���������Ƃ������Ă����Ă����v

