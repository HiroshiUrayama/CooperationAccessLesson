Attribute VB_Name = "Memo"
'#################################################
'Access�A�g�̃��b�X��
'accdb�t�@�C�����̃e�[�u���f�[�^�������ꍇ��DAO/ADO���g��(�ʁX�̋Z�p)
'#################################################
Option Explicit

'�y���̃f�[�^�x�[�X�Ɉڍs���l���Ă��Ȃ��zmdb(�G���f�B�[�r�[�t�@�C��)
'==================================================
'Dao=DataAccessObject
'�y�Q�Ɛݒ�FMicrosoft Office XX.X Access Database Engine Object Library(accdb�t�@�C��)�z
'�y�Q�Ɛݒ�FMicrosoft DAO 3.6 Object Library(mdb�t�@�C��)�^Access2003�܂ŕW��������Access�f�[�^�x�[�X�̌`���z
'Access�̏������痘�p����Ă���BAccess�f�[�^�x�[�X�̓����ɍł����v���Ă���B
'���̃f�[�^�x�[�X���g�p�����A�����I�ɂ�Access�݂̂Ƃ����ꍇ�ɗ��p����
'==================================================
'Database�I�u�W�F�N�g�F�J���Ă���f�[�^�x�[�X
'TableDef�I�u�W�F�N�g�F�ۑ�����Ă���e�[�u��
'QueryDef�I�u�W�F�N�g�F�ۑ�����Ă���N�G��
'Recordset�I�u�W�F�N�g�F�e�[�u����N�G���Œ�`����Ă��郌�R�[�h�̏W����

'---------------------------------------------------
'CreateObject�֐��Ő錾����ꍇ
'mdb�t�@�C��=CreateObject("DAO.DBEngine.36")
'accdb�t�@�C��=CreateObject("DAO.DBEngine.120")
'---------------------------------------------------

'---------------------------------------------------
'object.OpenRecordset(Name,Type,Options, LockEdit)
'---------------------------------------------------
'Object��Database�I�u�W�F�N�g���w�肷��
'�V����Recordset�I�u�W�F�N�g���쐬����Recordsets�R���N�V�����ɒǉ�����B
'����name�cRecordset�̃��R�[�h�̎擾�����w�肷��B
'�e�[�u�����E�N�G�����E���R�[�h��Ԃ�SQL�X�e�[�g�����g���w��ł���B
'����LockEdit�ŁARecordset�̃��b�N���w�肷��B
'��LockType

'���@dbOpenDynaset(2)�^�_�C�i�Z�b�g�^�C�v��Recordset���J���B
'���[�J���e�[�u���A�����N�e�[�u���A�I���N�G������쐬�ł���B���R�[�h�ǉ��A�X�V���\�B
'Recordset�̃��R�[�h���ύX�����ƁA���̃e�[�u���̃��R�[�h���X�V�����B

'���AdbOpenForwordOnly(8)�^�O���X�N���[���^�C�v��Recordset���J���B���R�[�h��O  �������݈̂ړ��\
'��Move���\�b�h��MoveNext���\�b�h�̂ݎg�p�\�B���R�[�h�Z�b�g�͍X�V�ł��Ȃ��B�@�\�͏��Ȃ��������ŏ����\�B

'���BdbOpenSnapshot(4)�^�X�i�b�v�V���b�g�^�C�v�̃��R�[�h�Z�b�g���J���B�X�i�b�v�V���b�g�쐬���̃f�[�^�ɌŒ肳��A
'���R�[�h�Z�b�g���X�V���邱�Ƃ��ł��Ȃ��

'���C�e�[�u���^�C�v��Recordset���J���B�e�[�u���ɓo�^����Ă��郌�R�[�h�̏W�܂��Ԃ��B���R�[�h�̒ǉ��E�폜�E�X�VOK�B
'�������A�����N�e�[�u���⌋���ɂ��I���N�G���ō쐬���邱�Ƃ͕s�\�B
'�܂��A���ɂȂ��Ă���e�[�u���ɍ쐬���ꂽ�u�C���f�b�N�X�v���g�p�ł���B
'���f�[�^�\�[�g�A�����������ɏ����ł���

'---------------------------------------------------
'����Options
'---------------------------------------------------
'�@dbAppendonly(8)�^'���[�U���V�������R�[�h���_�C�i�Z�b�g�ɒǉ�����̂������邪�A�����̃��R�[�h��ǂݎ�邱�Ƃ͋����Ȃ�
'�AdbConsistent(32)�^�_�C�i�Z�b�g���̑��̃��R�[�h�ɉe����^���Ȃ��t�B�[���h�ɂ̂ݍX�V��K�p����(�_�C�i�Z�b�g�^�C�v�ƃX�i�b�v�V���b�g�̂�)
'�BdbDenyRead(2)�^���̃��[�U�[��Recordset�̃��R�[�h��ǂݎ��Ȃ��悤�ɂ���(�e�[�u���^�C�v�̂�)
'�CdbDenyWrite(1)�^���̃��[�U�[��Recordset�̃��R�[�h��ύX�ł��Ȃ��悤�ɂ���
'�DdbForwardOnly(256)�^�O���X�N���[���݂̂̃X�i�b�v�V���b�g�^�C�vRecordset���쐬����(�X�i�b�v�V���b�g�^�C�v�̂�)
'�EdbInconsistent(16)�^���̃��R�[�h�ɉe�����y�ԏꍇ�ł��A�S�Ẵ_�C�i�Z�b�g�t�B�[���h�ɍX�V��K�p����(�_�C�i�Z�b�g�^�C�v�ƃX�i�b�v�V���b�g�^�C�v�̂�)
'�FdbReaOnly(4)�^Recordset��ǂݎ���p�Ƃ��ĊJ��

'---------------------------------------------------
'����LockEdit�Ɏw�肷��LockTypeEnum�񋓌^�̒l
'---------------------------------------------------
'�@dbOptimistic�^���R�[�hID�Ɋ�Â����L�I�������b�N�B
'�J�[�\���͌Â����R�[�h�ƐV�������R�[�hID���r���A���̃��R�[�h�ւ̃A�N�Z�X���Ō�ɍs���Ă���ύX��������ꂽ���ǂ������f����

'�AdbPessimistic�^�r���I�������b�N�B
'�J�[�\���ͤ���R�[�h���X�V�\�ł��邱�Ƃ�ۏ؂��邽�߂ɕK�v�ȍŒ���̃��b�N���g�p����

'�y���̃f�[�^�x�[�X�Ɉڍs����\��������zaccdb(�G�[�V�[�V�[�f�B�[�r�[�t�@�C��)
'==================================================
'ADO=ActiveXDataObject
'�y�Q�Ɛݒ�FMicrosoft ActiveX Data Object X.X Library�z
'�uOLEDB�v���o�C�_�v�Ƃ����d�g�݂���āAAccess�f�[�^�x�[�X�͂������A
'SQLServer��Oracle�Ȃǂ��������Ƃ��\�B�����I�ɑ��̃f�[�^�x�[�X�ɈڐA����\���̂���ꍇ�Ɏg�p����B
'��ADO�̋@�\���g������ADOX���g�p�����B
'���Q�Ɛݒ�FADO Ext. X.X for DDL and Security�ɎQ�Ɛݒ�
'��CreateObject�֐��̈����ɁuADOX.Catalog(�f�[�^�x�[�X�ɐڑ�����ꍇ)���w�肷��B
'��mdb�t�@�C���ƈႤ�Ƃ���
    '�@���Α��̃����[�V�����V�b�v�̓���
    '�A�t�B�[���h�̓��e�ɊO���t�@�C����Y�t���邱�Ƃ��ł���Y�t�t�@�C���^
    '�BSharePoint�AOutlook�Ƃ̘A�g����
    '�C����ȕ�������i�[�\�ȃ����^�ɂ����闚���Ǘ��@�\
'==================================================

'---------------------------------------------------
'CreateObject�֐����g��=ADO��Connection�I�u�W�F�N�g�FCreateObject("ADODB.Connection")
'CreateObject�֐����g��=ADO��Recordset�I�u�W�F�N�g�FCreateObject�i�hADODB.Recordset")
'---------------------------------------------------

'---------------------------------------------------
'����^�f�[�^�x�[�X���J���������Ώۂ̃f�[�^���擾���遨�f�[�^���������遨�f�[�^�x�[�X�����
'---------------------------------------------------

'ADO�̏ꍇ�ͤ�ڑ��������������Ɖ�������
'�@Provider�^�y�Q�Ɛݒ�FMicrosoft.ACE.OLEDB.XX.X;
    'accdb�t�@�C���̏ꍇ�BXX.X�ɂ̓o�[�W������\������������BAccess2016�́A16.0�B
    'mdb�t�@�C���̏ꍇ�́uMicrosoft.Jet.OLEDB.4.0;�v
    'SQLServer�̏ꍇ��SQLOLEDB���w�肷��
'�ADataSource
    '�t�@�C���p�X���w�肷��B
    
'�I�u�W�F�N�g���f��(ADO�̎�ȃI�u�W�F�N�g)
    '�@Connection�^�f�[�^�x�[�X�ւ̐ڑ���\���I�u�W�F�N�g
    '�ACommand�^�f�[�^�x�[�X�ɑ΂��Ď��s����R�}���h��ێ�����I�u�W�F�N�g�B
        '�N�G����SQL�X�e�[�g�����g�����s���鎞�Ɏg�p����B
    '�BRecordset�^�e�[�u����N�G���Œ�`����Ă��郌�R�[�h�̏W��
    


'������������������������������������������������������������
'��{�I�ɍ��݂����Ȃ����A���݂����ď����ꍇ������
'������������������������������������������������������������
'�Q�Ɛݒ�̗D�揇�ʂ���̂����ɓǂ܂��

'�ǂ������킩��₷�����邽�߂ɂ�
'---------------------------------------------------
'Dim rs As DAO.Recordset
'Dim rs As ADODB.Recordset
'DAO�Ȃ̂��AADODB�Ȃ̂���Recordset�̑O���ɂ�����Ə�������(���Ȃ݂ɁA�ȗ��\�炵��)

'�z���g��A���������
'Dim rs As Recordset

'�Q�Ɛݒ���s���Ă���ꍇ
'---------------------------------------------------
'�������ADO�Ȃ̂�DAO�Ȃ̂�������

'�Q�Ɛݒ���s���Ă��Ȃ��āA���s���o�C���f�B���O�̏ꍇ
'---------------------------------------------------
'�ϐ�����adoRS�Ȃ̂�daoRS�Ȃ̂��𖾋L����Ɨǂ�



























































