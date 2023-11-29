# Tracking DHL

## Descrição

Este script em Python foi desenvolvido para realizar consultas de rastreamento de remessas utilizando a API da DHL. O usuário pode fornecer um arquivo Excel com a extensão .xlsx contendo o cabeçalho "AWB" e os respectivos números de AWB , e o script retornará as informações de rastreamento associadas a esses números. 

## Dependências
- Python >= 3.0.0

Certifique-se de ter as bibliotecas listadas no arquivo requirements.txt do Python instaladas:

Você pode instalar as dependências utilizando o seguinte comando:
```bash
pip install -r requirements.txt
```

## Utilização

- Execute o script Python.
```bash
python main.py
```
- Selecione o arquivo Excel contendo os números de AWB.
- Forneça o Site ID do cliente e a senha quando solicitado.
- Escolha se deseja obter o histórico completo ou apenas o último evento.
- Aguarde o término da execução. O resultado será salvo em um novo arquivo Excel.


## Observações
O script utiliza a API XML da DHL para consultar informações de rastreamento.
Caso a requisição falhe, o script exibirá detalhes do erro, incluindo a resposta XML da DHL.
É possível personalizar o formato de saída ajustando a lógica dentro da função consultar_awb.
Certifique-se de tratar as informações de login com confidencialidade.
Observação: Certifique-se de estar em conformidade com os termos de uso da API da DHL ao utilizar este script. Esteja ciente de eventuais limitações e restrições impostas pela DHL em relação ao uso da API.