object Form1: TForm1
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  ClientHeight = 233
  ClientWidth = 628
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object lb1: TLabel
    Left = 24
    Top = 40
    Width = 107
    Height = 15
    Caption = 'Text to send to  "C2"'
  end
  object edText: TEdit
    Left = 24
    Top = 61
    Width = 585
    Height = 23
    TabOrder = 0
    Text = 'Teguh Prasetyo'
  end
  object btnSendToNewFile: TButton
    Left = 24
    Top = 90
    Width = 145
    Height = 25
    Caption = 'Send To New File'
    TabOrder = 1
    OnClick = btnSendToNewFileClick
  end
  object edSource: TEdit
    Left = 24
    Top = 144
    Width = 585
    Height = 23
    TabOrder = 2
    Text = 'edSource'
  end
  object edTarget: TEdit
    Left = 24
    Top = 173
    Width = 585
    Height = 23
    TabOrder = 3
    Text = 'Edit1'
  end
end
