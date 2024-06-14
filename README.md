markdown
Copiar código
# Projeto de Acompanhamento de Retirada de Rastreadores de Veículos de Locadoras

Este projeto tem como objetivo automatizar o processo de acompanhamento da retirada de rastreadores de veículos de locadoras. Através de uma série de scripts, o sistema verifica o status dos rastreadores, gera ordens de serviço (OS) quando necessário, e envia emails de questionamento aos fornecedores sobre a execução do serviço.

## Descrição dos Arquivos

### Tratar_Bases
- **Descrição**: Responsável por baixar as bases de dados necessárias do sistema e organizar as informações para uso posterior.
- **Função Principal**: Automatiza o download e a organização dos dados de veículos e filiais.

### Filiais
- **Descrição**: Verifica e valida as filiais onde os veículos se encontram.
- **Função Principal**: Confere as filiais atualizadas para garantir que as informações de localização dos veículos estejam corretas.

### Email
- **Descrição**: Compila informações e envia emails aos fornecedores para questionar se o serviço de retirada do rastreador já foi executado.
- **Função Principal**: Gera e envia emails de acompanhamento aos fornecedores com base no status das ordens de serviço.

### Gerar_Ordem
- **Descrição**: Realiza todo o processo para gerar a ordem de serviço de retirada do rastreador.
- **Função Principal**: 
  1. **Verificação de Data**: Verifica se a data de desativação do rastreador está vazia.
  2. **Verificação de Ordens Abertas**: Se a data estiver vazia, verifica se já existe uma ordem aberta no sistema.
  3. **Envio de Email de Questionamento**: Se houver uma ordem aberta, questiona o fornecedor sobre o motivo da não retirada.
  4. **Geração de Ordem de Serviço**: Se não houver ordem aberta, cria uma nova OS e envia um email ao fornecedor informando que o serviço precisa ser executado.
  5. **Validações**: Realiza validações no sistema para garantir que a ordem seja aberta corretamente, incluindo a verificação de nome, data e valores acordados em contrato.

### Processo Diário
- **Descrição**: O sistema baixa o relatório diariamente e repete todo o processo descrito acima.
- **Parada de Emails de Questionamento**: O envio de emails de questionamento para quando o fornecedor responde informando que o rastreador já foi retirado e o analista adiciona a data de desativação no sistema.

## Fluxo do Processo

1. **Baixa e Organização de Dados**:
    - Utiliza `Tratar_Bases` para baixar e organizar as informações necessárias.
2. **Verificação das Filiais**:
    - Utiliza `Filiais` para conferir as localizações dos veículos.
3. **Geração e Validação de Ordens de Serviço**:
    - Utiliza `Gerar_Ordem` para verificar o status dos rastreadores, gerar ordens de serviço e realizar todas as validações necessárias.
4. **Envio de Emails**:
    - Utiliza `Email` para questionar fornecedores e acompanhar a execução dos serviços.
5. **Processo Diário**:
    - Repete o processo diariamente até a confirmação da retirada dos rastreadores pelos fornecedores.

