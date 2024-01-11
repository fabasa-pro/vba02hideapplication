## Ocultar o Microsoft Office

Veja aqui como ocultar o `Microsoft Word`, deixando apenas a janela principal do `Visual Basic` visível, para quem precisa programar um `Desktop Application` e não quer ver aquele documento `Word` aberto o tempo todo.

1. Clique duas vezes em **ThisDocument** para exibir a janela **ThisDocument(Código)**, deixe **Document_Open** com o seguinte código-fonte:

```VBA
Option Explicit

Private Sub Document_Open()

    Project.ThisDocument.Application.Visible = False    ' Ocultar ou mostrar aplicativos.
    Project.UserForm1.Show 1                            ' Exibir como caixa de diálogo modal (restrita).
    
End Sub
```

2. Clique duas vezes em **UserForm1** para exibir o formulário e pressione F7 ou clique duas vezes no corpo do formulário para exibir a janela **UserForm1(Código)** e deixe com o seguinte código-fonte:

```VBA
Option Explicit

Private Sub UserForm_Terminate()

    Project.ThisDocument.Application.Visible = True                                                    ' Ocultar ou mostrar aplicativos.
    Project.ThisDocument.Application.Quit SaveChanges:=wdSaveChanges, OriginalFormat:=wdWordDocument   ' Salvar e fechar tudo.

End Sub

```

> :bell: **Importante:** <br> Se você adicionou este código é impossível retornar ao `Word` e para isso você precisa manter pressionada a tecla Shift e executar com Shift pressionada. Desta forma você pode ignorar todo o `VBA` e entrar novamente no aplicativo `Word` para programar.

![screenshot](https://github.com/fabasa-pro/vba02hideapplication/blob/main/vba02hideapplication.png)

Veja como aplicar [aplicar propriedades de formulário](https://github.com/fabasa-pro/vba03formborderstyle) para **Minimizar**, **Maximizar** e **Restaurar** a janela e aplicar a propriedade **WindowState** para exibir diferentes tipos de **Bordas** de formulário.

## Licenciado sob a licença MIT

Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.

Consulte [LICENSE.TXT](https://github.com/fabasa-pro/vba01userform/blob/main/LICENSE.TXT) na raiz do projeto para obter informações.
