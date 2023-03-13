Attribute VB_Name = "backup"
Option Explicit

'=======================================================
'�y�T  �v�z�o�b�N�A�b�v�쐬�̃��C������
'�y��  ���z������    ����
'          ---------------------------------------------
'          �Ȃ�
'          ---------------------------------------------
'�y�߂�l�z�Ȃ�
'�y��  �l�z�Ȃ�
'=======================================================
Public Sub save()

  Dim book As Workbook, path As String
  Set book = ActiveWorkbook
  
  '�o�b�N�A�b�v�t�H���_�̑��݊m�F/�쐬
  path = book.path & "\" & Split(book.Name, ".")(0) & "_backup"
  makeFolder path
  
  '�T�u�t�H���_(���t����)�̑��݊m�F/�쐬
  path = path & "\" & Format(Date, "yyyymmdd")
  makeFolder path
  
  '�o�b�N�A�b�v�̍쐬
  path = path & "\" & getNameWithTimeStamp(book)
  book.SaveCopyAs path
  
End Sub

'=======================================================
'�y�T  �v�z�t�H���_�쐬
'�y��  ���z������    ����
'          ---------------------------------------------
'          path      String�^�A�t�H���_�p�X
'          ---------------------------------------------
'�y�߂�l�z�Ȃ�
'�y��  �l�z�Ȃ�
'=======================================================
Private Sub makeFolder(path As String)
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  '�Ώۃt�H���_�̑��݊m�F(����: �����Ȃ��A�Ȃ�: �쐬)
  If Not fso.FolderExists(path) Then
      fso.CreateFolder path
  End If
End Sub

'=======================================================
'�y�T  �v�z�t�@�C�����ɌĂяo�����̎��Ԃ�t�^
'�y��  ���z������    ����
'          ---------------------------------------------
'          book      Workbook�^
'          ---------------------------------------------
'�y�߂�l�z�Ăяo�����̎��Ԃ�t�^�����t�@�C����
'�y��  �l�z�}�N���̗L��/�����Ŋg���q��I�����Ă���
'=======================================================
Private Function getNameWithTimeStamp(book As Workbook) As String

  Dim baseName As String, suffix As String, extension As String
  
  '���O�E�ڔ����E�g���q�̗p��
  baseName = Split(book.Name, ".")(0)
  suffix = Format(Now, "_hhnn")
  extension = IIf(book.HasVBProject, ".xlsm", ".xlsx")
  
  getNameWithTimeStamp = baseName & suffix & extension
End Function
