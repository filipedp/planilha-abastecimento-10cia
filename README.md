# planilha-abastecimento-10cia
Automação de controle de abastecimentos da 10ª Companhia Independente da PMES, utilizando Google Apps Script.
# 🚓 Automação de Abastecimentos em Google Sheets

Este projeto automatiza o controle de abastecimentos de veículos utilizando Google Apps Script integrado com Google Sheets, Gmail e uma API externa de abastecimento. Ele foi desenvolvido para facilitar o gerenciamento de notas fiscais, cruzamento de dados e envio de notificações automáticas.

---

## 📌 Funcionalidades

✅ Leitura automática de arquivos XML de Notas Fiscais recebidos por e-mail  
✅ Integração com API externa para importação de dados de abastecimento em lote  
✅ Preenchimento automático da planilha de controle com dados dos abastecimentos  
✅ Verificação de correspondência entre dados de abastecimento e notas fiscais  
✅ Envio de e-mails automáticos de confirmação ou solicitação de regularização  
✅ Menu personalizado no Google Sheets para execução manual das funções principais

---

## 🧠 Estrutura do Código

- `parseNewEmailsAndPopulateSheet()`  
  Lê e-mails não lidos com arquivos XML em anexo e extrai os dados relevantes.

- `buscarAbastecimentosSisatecXML()`  
  Consulta uma API externa e importa os dados de abastecimento em formato XML.

- `atualizarNotas()`  
  Atualiza a planilha com base nos dados coletados, preenchendo número da nota e série.

- `enviarConfirmacoesDeNota()`  
  Envia e-mails de confirmação de recebimento ou solicita preenchimento de formulário.

- `onOpen()`  
  Cria um menu personalizado para execução manual das funções principais diretamente pela interface da planilha.

---

## 📂 Pré-requisitos

- Conta Google com acesso autorizado às planilhas e permissões para:
  - Ler e-mails do Gmail
  - Executar scripts no Google Sheets
- Estrutura mínima da planilha:
  - Aba de controle de abastecimentos
  - Aba para armazenamento de e-mails processados
  - Aba com cadastro de notas previamente registradas
  - Aba com cadastro de destinatários (nomes, e-mails, etc.)

---

## 🔗 Formulário de Justificativa

Nos casos em que a nota fiscal não é localizada automaticamente, um e-mail é enviado ao responsável com um link para preenchimento de um formulário de regularização (formulário configurado pelo administrador da planilha).

---

## 🛠️ Tecnologias Utilizadas

- Google Apps Script (JavaScript)
- Google Sheets
- Gmail API
- API externa de abastecimento (XML)

---

## 📄 Licença

Este projeto é distribuído sob a Licença MIT.  
Sinta-se à vontade para utilizar, adaptar e contribuir.

---

## 🙋‍♂️ Observações

Este projeto foi desenvolvido com foco em automação administrativa, sendo adaptável a qualquer organização que utilize planilhas para controle de abastecimento ou gestão de notas fiscais.
