# Oninmetro

![Oninmetro](./images/oninmetro.png)

A ideia desse projeto é criar um robô simples que realiza a tarefa de enviar um email semanal de solicitação de Transporte para o Inmetro e automatizar esse envio toda quinta-feira, entre 13h00 e 14h00, utilizando duas ferramentas: **GithubActions** e **Google App Scripts**.
> **Nota:** A plataforma de email é o [Webmail](https://webmail.inmetro.gov.br/owa/). 

- Primeiro, os passeiros preenchem um formulário, que pode ser acessado [AQUI](https://forms.gle/Pj3YoQni55LQSXZ66), selecionando seus dias para uso do transporte para a **SEMANA SEGUINTE**. 
- ⚠️Esse código coleta os dados preenchidos entre **14:30** da **ULTIMA QUINTA FEIRA** até o momento da sua execução.
- É permissível os passageiros responderem mais de uma vez o formulário, caso mudem de ideia. Será considerado apenas a última resposta.
- O email é enviado automaticamente toda **Quinta-feira por volta das 13:00 (BRT)** via GitHub Actions, disparado por um agendamento no Google Apps Script. Se alguém responder após esse horário, sua solicitação será desconsiderada para a semana seguinte, pois o email já foi enviado. Ele deverá solicitar o transporte por si mesmo.
- Voce pode ver as respostas do formulário [AQUI](https://docs.google.com/spreadsheets/d/1VeInFHfXwIrWc06CSn6ode2IRsThyWITMrqFIV1cXoU/edit?usp=sharing).

## RUN LOCAL
Siga estes passos a seguir se você quer testar o robô localmente na sua máquina.
1. Se já tem o `Python` na sua máquina, vai precisar das seguintes instalações:
```bash
pip install pandas, selenium, webdriver_manager, datetime, time, traceback
```
2. Agora clone esse repositório:
```bash
git clone https://github.com/josewilsonsouza/oninmetro.git
```
3. Navegue até a pasta `oninmetro` e rode o script `mail.py`
```bash
python mail.py
```
> Para interromper o robô, basta apertar `Ctrl + C` no terminal.
4. Acompenhe no terminal, ele pedirá um usuário e senha do Webmail.

<img width="499" height="83" alt="image" src="https://github.com/user-attachments/assets/1b2a3ea4-50c0-437a-8b5b-7d5703ee9cad" />

5. Depois disso, pode ocorrer de o formulário não ter respostas registradas desde às 14:30 da última quinta-feira. Então, você pode responda ao [FORMS](https://forms.gle/Pj3YoQni55LQSXZ66) para fazer os testes.
6. Ele pedirá para confirmar se a mensagem do email está correta, digite `s`

  <img width="428" height="107" alt="image" src="https://github.com/user-attachments/assets/7237fadc-8246-477e-9c9a-f3703f985473" />
   
7. Pronto, agora veja as janelas abrindo e o robô trabalhando. No arquivo `mail.py` você pode editar as informações:
   - Campo `Para`: `destinatario@inmetro.gov.br`
   - Campo `Cc`: `copia1@inmetro.gov.br; copia2@inmetro.gov.br; copia3@inmetro.gov.br`
   - Campo `Assunto`: `Solicitação de utililização do transporte Inmetro`
   - Também pode editar alguns outros detalhes como seu nome (depois do Atensiosamente).
   - Corpo do email: fará uma cópia das informações dos passageiros.

8. Depois pode digitar `s` para enviar.

## ENVIO AUTOMÁTICO

O envio automático funciona com duas peças:

1. **Google Apps Script** - Um script externo agendado para toda quinta-feira ~13:00 (BRT) que dispara o workflow do GitHub Actions via API (`workflow_dispatch`).
2. **GitHub Actions** - O workflow (`.github/workflows/send-email.yml`) executa o `mail.py --auto` em um runner Ubuntu com Chrome e xvfb. As credenciais e destinatários são configurados via **GitHub Secrets**:
   - `OWA_USERNAME` - Usuário do Webmail
   - `OWA_PASSWORD` - Senha do Webmail
   - `EMAIL_DESTINO` - Endereço de destino
   - `EMAILS_COPIA` - Endereços em cópia (separados por `;`)

O workflow também pode ser disparado manualmente pela aba **Actions** no GitHub (botão "Run workflow").
