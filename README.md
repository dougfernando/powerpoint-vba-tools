# DFS Tools para PowerPoint

Add-in VBA (`.ppam`) para executar ferramentas de texto, formas, tabelas, QA e integração com Excel a partir da aba **DFS Tools** do PowerPoint.

O projeto mantém as fontes VBA, o catálogo de comandos, a interface do launcher e o Ribbon fora do arquivo binário. O PPAM é reconstruído por scripts PowerShell, tornando as mudanças revisáveis e reproduzíveis.

## Experiência do launcher

Use **DFS Tools > Abrir macros**. O formulário é modeless: ele pode permanecer aberto enquanto você troca de slide, seleção ou apresentação.

- A lista mostra várias macros no formato `Nome da macro (Categoria)`.
- O foco inicial fica na lista de macros.
- Digitar o começo do nome seleciona uma macro na lista.
- `Enter` executa a macro selecionada e `Esc` fecha o launcher.
- `Alt+C`, `Alt+M`, `Alt+E` e `Alt+F` acessam categoria, macros, Executar e Fechar.
- Duplo clique também executa uma macro.
- A área **Resultado** é o único canal de mensagens do add-in; não são usados `MsgBox`.
- O resultado desaparece automaticamente depois de 1 segundo.

Operações destrutivas usam confirmação não modal. A primeira execução apresenta o pedido na área **Resultado**; execute o mesmo comando novamente dentro de 1 segundo para confirmar.

O catálogo atual contém 15 comandos nas categorias Texto, Formas, Tabelas, QA e Excel. A fonte oficial do catálogo é [src/commands.json](src/commands.json).

## Requisitos

- Windows com PowerPoint desktop.
- Windows PowerShell 5.1.
- Excel desktop apenas para os comandos de exportação e importação de textos.
- Permissão temporária para **Confiar no acesso ao modelo de objeto do projeto VBA** durante build, importação, exportação ou recarga de módulos.

Essa opção fica em **Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Configurações de Macro**. Políticas corporativas podem bloqueá-la.

O uso normal do add-in não exige acesso ao projeto VBA. Os scripts não habilitam macros e não alteram permanentemente políticas de segurança do Office.

## Início rápido

Salve o trabalho aberto e feche todas as instâncias do PowerPoint. Na raiz do projeto, execute:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-ppam.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-addin.ps1
```

Abra o PowerPoint e acesse **DFS Tools > Abrir macros**.

O artefato principal é gerado em:

```text
build\DouglasPowerPointTools.ppam
```

O instalador copia o add-in para:

```text
%APPDATA%\Microsoft\AddIns\DouglasPowerPointTools\DouglasPowerPointTools.ppam
```

Se houver uma instalação antiga registrada em outro caminho, remova-a em **Arquivo > Opções > Suplementos > Gerenciar: Suplementos do PowerPoint > Ir** antes de instalar novamente.

## Desenvolvimento e recarga rápida

O botão **DFS Tools > Recarregar macros** relê os módulos `.bas` gerenciados em `src`. O carregador:

1. localiza o projeto VBA do PPAM;
2. cria backup dos módulos gerenciados;
3. fecha o launcher;
4. substitui os módulos;
5. restaura o backup se a importação falhar;
6. reabre o launcher e mostra o resultado na própria interface.

O caminho de `src` usado no build é incorporado ao PPAM. Se ele não existir, o add-in solicita uma pasta e salva essa escolha nas configurações do usuário.

A recarga rápida serve para alterações em módulos `.bas` gerenciados. As mudanças abaixo exigem novo build:

- [src/commands.json](src/commands.json);
- [src/ui/launcher.json](src/ui/launcher.json);
- [src/ui/frmMacroLauncher.code](src/ui/frmMacroLauncher.code);
- [resources/customUI.xml](resources/customUI.xml).

## Estrutura do projeto

| Caminho | Responsabilidade |
| --- | --- |
| `src/modCommands_*.bas` | Comandos públicos consumidos pelo launcher |
| `src/modNotifications.bas` | Mensagens, limpeza temporizada e confirmação não modal |
| `src/modMacroLauncher.bas` | Estado e comportamento geral do launcher |
| `src/modLoader.bas` | Recarga segura dos módulos durante desenvolvimento |
| `src/modPresentation.bas` | Resolução da apresentação e do slide ativos |
| `src/modSelection.bas` | Validação da seleção do PowerPoint |
| `src/mod*.bas` | Utilitários compartilhados de texto, formas e configuração |
| `src/commands.json` | Catálogo explícito de comandos |
| `src/ui/launcher.json` | Controles e geometria do UserForm |
| `src/ui/frmMacroLauncher.code` | Eventos, teclado e aparência do UserForm |
| `resources/customUI.xml` | Aba DFS Tools no Ribbon |
| `scripts/lib/` | Modelo de fontes, empacotamento e utilitários COM/Open XML |
| `scripts/` | Build, instalação, desinstalação e sincronização com o VBE |
| `tests/test-offline.ps1` | Testes estruturais sem abrir o Office |
| `docs/VALIDATION.md` | Roteiro de validação manual no PowerPoint |
| `build/` | Artefatos, staging, backups e logs; não versionar |

`modCommandCatalog` e `frmMacroLauncher` são gerados durante o build. Não os edite dentro de um PPAM como fonte definitiva.

## Adicionar ou alterar uma macro

Cada comando do launcher deve:

1. ser uma `Public Sub Cmd_...()` sem parâmetros em um módulo de `src`;
2. possuir uma entrada correspondente em `src/commands.json`;
3. usar chamadas diretas qualificadas pelo módulo, geradas no catálogo;
4. publicar sucesso, validação, cancelamento ou erro com `Notify`;
5. não usar `MsgBox`;
6. usar `ConfirmThroughStatus` antes de uma exclusão que exija confirmação.

Exemplo mínimo:

```vb
Public Sub Cmd_Example_DoSomething()
    On Error GoTo ErrorHandler

    ' Operação...
    Notify "Operacao concluida."
    Exit Sub

ErrorHandler:
    Notify "Falha na operacao: " & Err.Description
End Sub
```

Depois da alteração, valide as fontes. Se apenas módulos `.bas` mudaram, use **Recarregar macros**. Se catálogo, formulário ou Ribbon mudaram, gere e instale um novo PPAM.

## Testes e validação

Validação rápida, sem abrir o Office:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-offline.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-ppam.ps1 -ValidateOnly
```

Esses testes verificam, entre outros contratos:

- sintaxe dos scripts PowerShell;
- catálogo e assinaturas `Cmd_*`;
- ausência de `MsgBox` no runtime;
- controles e dimensões do launcher;
- atalhos de teclado e foco inicial;
- timer de resultado;
- geração do dispatcher direto;
- estrutura Open XML do PPAM;
- publicação com backup e proteção contra arquivo bloqueado.

Os testes offline não substituem a compilação VBA e a verificação visual no PowerPoint. Antes de distribuir, siga [docs/VALIDATION.md](docs/VALIDATION.md).

## Como o build funciona

O script [scripts/build-ppam.ps1](scripts/build-ppam.ps1):

1. valida fontes, catálogo e definição da interface;
2. abre uma instância isolada do PowerPoint;
3. importa os módulos VBA e constrói o UserForm;
4. gera `modCommandCatalog` com chamadas diretas;
5. salva o projeto com macros;
6. tenta salvar no formato PPAM nativo (`30`);
7. quando necessário, converte o content type do PPTM para PPAM preservando `vbaProject.bin`;
8. incorpora o RibbonX;
9. valida o pacote e, por padrão, executa o health check;
10. publica atomicamente o novo arquivo, mantendo backup do anterior.

Se somente a validação COM de carregamento falhar, é possível ignorar o health check:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-ppam.ps1 -SkipRuntimeValidation
```

Isso não ignora a validação estrutural do pacote. Instale e teste manualmente o launcher depois.

Os parâmetros `SourceFolder`, `OutputFolder` e `OutputFile` também estão disponíveis. `OutputFile` deve ser somente um nome terminado em `.ppam`, sem caminho.

## Importar e exportar pelo VBE

Para trabalhar em um PPTM de desenvolvimento já aberto, informe explicitamente seu caminho:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\import-src-into-active-presentation.ps1 -PresentationPath C:\Caminho\DouglasPowerPointTools.build.pptm
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\export-active-vba-to-src.ps1 -PresentationPath C:\Caminho\DouglasPowerPointTools.build.pptm
```

A importação cria uma cópia de segurança antes de substituir módulos e não salva automaticamente a apresentação alterada. A exportação atualiza somente módulos `.bas` gerenciados já existentes em `src`; novos componentes ficam na pasta intermediária para revisão.

As fontes do repositório usam UTF-8. Antes de importar no VBIDE, os scripts convertem os módulos para a página ANSI local e recusam caracteres que não possam ser representados.

## Desinstalação e recuperação

Para desinstalar:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\uninstall-addin.ps1
```

O registro é removido e o arquivo instalado é renomeado para `.uninstalled`, permitindo recuperação. Backups não são apagados automaticamente.

Em caso de falha de build, consulte a pasta `build\staging-*` indicada no terminal e seu `build-error.log`. Se o destino estiver bloqueado, feche o PowerPoint e confirme que não existe uma instância `POWERPNT` em segundo plano. O script não encerra sessões do usuário automaticamente.

Para instalar manualmente um backup ou outro PPAM:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-addin.ps1 -PpamPath C:\Caminho\DouglasPowerPointTools.ppam
```

## Referências

- [Roteiro de validação manual](docs/VALIDATION.md)
- [PowerPoint Presentation.SaveAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas)
- [PowerPoint AddIn.Registered](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.addin.registered)
- [PowerPoint Application.Run](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.run)
