VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Licenciado sob a licença MIT.
' Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.

' ==========================================================================
' NOTA: para editar o código-fonte, executar o arquivo com a tecla <Shift>
' pressionada para ignorar todo o VBA e entre no aplicativo Microsoft Word.
' ==========================================================================

Option Explicit

Private Sub UserForm_Terminate()

    Project.ThisDocument.Application.Visible = True                                                    ' Ocultar ou mostrar aplicativos.
    Project.ThisDocument.Application.Quit SaveChanges:=wdSaveChanges, OriginalFormat:=wdWordDocument   ' Salvar e fechar tudo.

End Sub
