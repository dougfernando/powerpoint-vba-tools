# Validação no PowerPoint

Execute em cópias descartáveis. O arquivo de testes offline usa apenas fixtures temporárias e não instala add-ins.

## Verificações automáticas

1. Execute tests/test-offline.ps1 em Windows PowerShell 5.1.
2. Execute build-ppam.ps1 -ValidateOnly.
3. Feche o PowerPoint e execute o build completo. A mensagem de sucesso exige gravação nativa, pacote válido, carregamento e health check do formulário.
4. O PPAM anterior deve permanecer disponível se o build falhar. Verifique backup, etapa e HRESULT no log.

Um health check não substitui a compilação completa: procedimentos ainda não chamados podem conter erros.

## Compilação e interface

1. Abra o PPTM intermediário indicado pelo build. No VBE, selecione seu projeto e execute Depurar > Compilar. Verifique referências ausentes e registre o resultado. Não altere outros projetos.
2. Feche o PowerPoint e instale o PPAM.
3. Abra duas apresentações descartáveis. A aba DFS Tools deve aparecer uma única vez.
4. Clique em Abrir macros repetidamente: deve existir uma única janela. Feche pelo botão e pelo X, e abra novamente.
5. Verifique as cinco categorias, a opção Todas, os 14 comandos e suas descrições.
6. Selecione formas com a janela aberta, troque de apresentação e execute. Somente a apresentação ativa deve mudar.
7. Sem apresentação ou com seleção inválida, espere uma mensagem compreensível, sem falha não tratada.
8. Crie em outra apresentação uma macro homônima com efeito claramente diferente; o launcher deve chamar exclusivamente a implementação do add-in.
9. Execute ao menos um caso de sucesso, um de validação e um erro. A label **Resultado** deve refletir cada desfecho em português, sem abrir `MsgBox`, e ficar vazia aproximadamente 1 segundo depois.
10. Em cada comando destrutivo, confirme que a primeira execução apenas solicita confirmação na label e que a segunda execução, dentro de 1 segundo, efetua a exclusão.
11. Confirme que a lista exibe vários comandos simultaneamente no formato `Nome (Categoria)`, aceita busca incremental pelo nome e responde a `Alt+C` e `Alt+M` independentemente do controle focado.
9. Um erro não pode provocar nova tentativa automática nem mostrar sucesso. Durante execução, o botão Executar deve ficar desabilitado.

## Comandos

| Comando | Cenário e resultado esperado |
| --- | --- |
| Desativar ajuste automático | Texto comum e dentro de grupos; ajuste automático desativado |
| Copiar texto para forma | Duas formas compatíveis; texto e formato copiados, origem excluída |
| Aplicar margens padrão | Formas selecionadas com texto; margens de 0,2 cm |
| Trocar textos | Duas formas selecionadas; conteúdos trocados |
| Remover quebras manuais | Um shape com Enter, Shift+Enter e quebras LF/CRLF; todas substituídas por espaços |
| Excluir pela cor | Referência preenchida; confirmar/cancelar; somente correspondências excluídas |
| Excluir semelhantes | Referência e duplicatas em slides diferentes; posição e tamanho considerados |
| Remover fora do slide | Formas totalmente fora e parcialmente dentro; somente as totalmente fora removidas |
| Converter tabela | Tabela selecionada; formas criadas e tabela original preservada |
| Mostrar IDs | Slide com placeholders; lista de identificadores |
| Marcar fontes | Fontes permitidas e não permitidas; somente as não permitidas marcadas |
| Limpar marcadores | Marcadores QA e formas normais; somente marcadores removidos |
| Verificar clientes | Texto com e sem termos configurados; correspondências marcadas |
| Exportar textos | Apresentação salva; XLSX criado ao lado; apresentação não salva deve ser recusada |
| Atualizar pelo Excel | Planilha válida e cancelamento; textos atualizados somente no caso válido |

Teste também erro de gravação e arquivo Excel inválido: não pode haver mensagem de sucesso após erro. A importação Excel mantém o comportamento existente por índice de slide e ID de forma; não reorganize slides entre exportação e atualização.

## Instalação e recuperação

- Instale, reinicie o PowerPoint e confirme o carregamento automático.
- Gere outra versão, reinstale e confira a substituição sem duplicar a aba.
- Valide backup e restauração usando cópias descartáveis, inclusive arquivo de destino bloqueado.
- Desinstale; a aba deve desaparecer após reinício, com o arquivo anterior preservado como .uninstalled.
- Com macros normalmente autorizadas, desative apenas o acesso ao modelo de objeto do projeto VBA e reinicie. Launcher e comandos devem continuar funcionando.
- Restaure a preferência de desenvolvimento se for gerar novos builds.

Não altere políticas corporativas para fazer o teste passar. Registre qualquer bloqueio como pendência.

## Registro desta implementação

Os resultados automáticos e eventuais impedimentos da máquina são informados na entrega. Não considerar itens manuais acima aprovados sem executá-los no Office.
