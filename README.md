# Gerador de Links do WhatsApp para Envio em Massa

## Descrição

Este script em **Python** automatiza a criação de **links personalizados do WhatsApp Web** para acelerar o processo de contato inicial com jovens selecionados para os cursos do IOS.

Ele lê uma planilha **Excel (`contatos.xlsx`)**, identifica um telefone válido para cada contato, gera uma **mensagem personalizada**, cria o **link do WhatsApp com a mensagem já preenchida** e exporta os resultados para uma nova planilha **(`links_wpp.xlsx`)**.

---

# Requisitos

Bibliotecas Python necessárias:

- pandas
- openpyxl

Instalação:

```bash
pip install pandas openpyxl