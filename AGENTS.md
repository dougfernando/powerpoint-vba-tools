# AGENTS.md

Este arquivo orienta agentes e automações que alterem este repositório. Ele se aplica a todo o projeto. Instruções mais específicas em um `AGENTS.md` de subdiretório, caso existam no futuro, prevalecem naquele escopo.

## Objetivo do projeto

Manter um add-in PowerPoint VBA reproduzível, chamado **DFS Tools**, cujo launcher apresenta comandos definidos em fontes versionáveis e gera um PPAM instalável sem depender de edição manual do binário.

## Fonte de verdade

- Edite arquivos em `src`, `resources`, `scripts`, `tests` e `docs`.
- Não trate arquivos de `build` ou componentes extraídos de um PPAM como fonte definitiva.
- `src/commands.json` é a fonte do catálogo.
- `src/ui/launcher.json` define controles e geometria do formulário.
- `src/ui/frmMacroLauncher.code` define eventos, teclado e aparência.
- `resources/customUI.xml` define o Ribbon.
- `modCommandCatalog` e `frmMacroLauncher` são gerados no build; não crie versões manuais concorrentes.

## Contratos do runtime

Preserve estes comportamentos, salvo pedido explícito em contrário. Quando um requisito mudar, atualize implementação, testes e documentação juntos.

- A aba do Ribbon se chama `DFS Tools`.
- O launcher é modeless e usa a apresentação/seleção ativas no momento da execução.
- A lista mostra comandos como `Nome da macro (Categoria)` e recebe o foco inicial.
- `Enter` executa, `Esc` fecha e os atalhos `Alt+C`, `Alt+M`, `Alt+E`, `Alt+F` funcionam independentemente do controle focado.
- A área `lblResult` é o único canal de mensagens do runtime.
- Não introduza `MsgBox` em módulos carregados nem no código do formulário.
- `InputBox` só é aceitável quando a macro precisa receber dados do usuário e não existe controle equivalente no launcher.
- Todo caminho final de um comando deve chamar `Notify` com mensagem curta e útil em português.
- O resultado é limpo de forma assíncrona; não use `Sleep`, loops de espera ou outra operação que bloqueie o PowerPoint.
- Operações destrutivas que exijam confirmação devem usar `ConfirmThroughStatus`, nunca diálogo modal.
- Ao fechar, trocar de macro ou iniciar outra execução, cancele timers e confirmações pendentes.

O prazo atual de status e confirmação é 1 segundo, definido e testado em `modNotifications.bas`. Não duplique esse valor em novos pontos sem necessidade.

## Contrato dos comandos

Um comando exposto deve:

1. ter nome `Cmd_*`;
2. ser `Public Sub` sem parâmetros;
3. residir em um módulo listado por `module` no catálogo;
4. possuir exatamente uma entrada em `src/commands.json`;
5. tratar erros e publicar o resultado com `Notify`;
6. usar helpers compartilhados de apresentação, seleção, texto ou formas quando disponíveis.

Não adicione comandos públicos `Cmd_*` fora do catálogo. Não use `Application.Run` para despachar comandos: o catálogo gerado utiliza chamadas diretas e qualificadas pelo módulo.

Ao adicionar um comando, atualize também a documentação e os cenários de validação manual quando o comportamento não estiver coberto pelos casos existentes.

## Regras para módulos e recarga

- Módulos `.bas` recarregáveis devem conter `@ManagedByDouglasTools` no cabeçalho.
- `modLoader` é deliberadamente estável e não deve receber essa marca.
- O carregador nunca deve remover a si próprio nem `modCommandCatalog`.
- Preserve backup e restauração transacional durante a recarga.
- Alterações somente em módulos `.bas` gerenciados podem ser testadas com **Recarregar macros**.
- Alterações no catálogo, formulário ou Ribbon exigem build completo.
- Evite dependências em `ThisPresentation`, que não é um global confiável do PowerPoint VBA.
- Não acesse o projeto do add-in pela coleção `Presentations`; use a resolução via projetos do VBE já implementada no carregador.

## VBA e compatibilidade

- Mantenha compatibilidade com Office VBA de 32 e 64 bits.
- Declarações de API devem usar compilação condicional e `PtrSafe`/`LongPtr` sob `VBA7`.
- Não adicione referências externas quando late binding resolver o caso.
- Fontes são UTF-8; o build converte para a página ANSI local antes da importação no VBIDE.
- Prefira mensagens em português e mantenha-as curtas o suficiente para `lblResult`.
- Preserve `Option Explicit` em todos os módulos e no formulário.
- Qualifique tipos do Office quando isso eliminar ambiguidade sem exigir uma nova referência.

## PowerShell e build

- Scripts precisam funcionar no Windows PowerShell 5.1.
- Evite recursos exclusivos do PowerShell 7.
- Libere objetos COM explicitamente nos fluxos que criam PowerPoint, apresentações, projetos ou componentes.
- Nunca encerre sessões do PowerPoint pertencentes ao usuário.
- Antes de build/publicação, detecte `POWERPNT` aberto e instrua o usuário a fechá-lo.
- Preserve publicação atômica e backup do PPAM anterior.
- Preserve `vbaProject.bin` ao usar a conversão Open XML de PPTM para PPAM.
- Não reduza validações do pacote para contornar erros de ambiente; use `-SkipRuntimeValidation` somente para ignorar o health check COM documentado.
- Não grave segredos, caminhos pessoais fixos ou políticas de segurança no repositório.

## Interface do launcher

- Adicione tipos de controle aceitos também à validação em `scripts/lib/common.ps1`.
- Atualize a lista de controles obrigatórios quando um controle estrutural for renomeado.
- Mantenha geometria suficiente para DPI do Windows; reserve margem inferior entre o último controle e a altura do formulário.
- Ao renomear controles, atualize `modMacroLauncher.bas`, o código do formulário e os testes.
- Use eventos explícitos de teclado para atalhos críticos; não dependa apenas de `Accelerator` do MSForms.
- O foco automático deve ocorrer apenas ao abrir o launcher, sem roubar foco durante a interação posterior.

## Testes obrigatórios

Após qualquer alteração de código ou configuração, execute:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-offline.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-ppam.ps1 -ValidateOnly
```

Para alterações em VBA, formulário, catálogo, Ribbon, empacotamento ou instalação, faça também o build completo quando o ambiente tiver PowerPoint disponível e fechado:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-ppam.ps1
```

Não declare compilação VBA ou validação visual como concluída apenas com testes offline. Informe claramente o que foi automatizado e o que ainda exige PowerPoint real. Use `docs/VALIDATION.md` como roteiro manual.

## Definition of done

Uma evolução está concluída quando:

- fontes e catálogo estão coerentes;
- mensagens e caminhos de erro seguem o contrato de `Notify`;
- interface e atalhos continuam acessíveis por teclado;
- testes offline e `-ValidateOnly` passam;
- README e `docs/VALIDATION.md` refletem mudanças visíveis ou operacionais;
- nenhum artefato de `build`, backup, staging ou arquivo temporário foi incluído como fonte;
- limitações de validação no Office foram comunicadas na entrega.
