# RoboMail Inmetro

A ideia desse projeto é criar um robô simples que realiza a tarefa de enviar um email semanal para o destinatário: `destinatario@inmetro.gov.br`. A plataforma de email é o [Webmail](https://webmail.inmetro.gov.br/owa/).

- Primeio, uma enquete é enviada na semana atual, X, que pode ser acessada [AQUI](https://forms.gle/Pj3YoQni55LQSXZ66). Essa enquete pode ser preenchida de segunda à quinta, de X, até às 14:29H, aceita modificações nesse período.
- Chegando na quinta (ou quando esse código for rodado), automaticamente as respostas serão detectadas, e as informações de email preenchidas.

## COMO USAR
1. Se já tem o `Python` na sua máquina, vai precisar das seguintes instalações:
```bash
pip install pandas, selenium, webdriver_manager, datetime, time, traceback
```
2. Agora clone esse repositório:
```bash
git clone https://github.com/josewilsonsouza/robomail.git
```
3. Navegue até a pasta `robomail` e rode o script `mail.py`
```bash
python mail.py
```
4. Acompenhe no terminal, ele pedirá um usuário e senha do Webmail.

<img width="499" height="83" alt="image" src="https://github.com/user-attachments/assets/1b2a3ea4-50c0-437a-8b5b-7d5703ee9cad" />

5. Depois disso, pode ocorrer de o formulário não ter respostas registradas na semana atual, X (pois é teste). Você pode acessar a planilha de **RESPOSTAS**
   [AQUI](https://docs.google.com/spreadsheets/d/1VeInFHfXwIrWc06CSn6ode2IRsThyWITMrqFIV1cXoU/edit?usp=sharing), edite manualmente algumas datas para fazer o teste. Mas se você está rodando dentro da janela de seg-qui, pode criar respostas no formulário, diretamente.
6. Ele pedirá para confirmar se a mensagem está correta, digite `s`

  <img width="428" height="107" alt="image" src="https://github.com/user-attachments/assets/7237fadc-8246-477e-9c9a-f3703f985473" />
   
7. Pronto, agora veja as janelas abrindo e o robô trabalhando. No arquivo `mail.py` você pode editar as informações:
   - Campo `Para`: `destinatario@inmetro.gov.br`
   - Campo `Cc`: `copia1@inmetro.gov.br; copia2@inmetro.gov.br; copia3@inmetro.gov.br`
   - Campo `Assunto`: `Solicitação de utililização do transporte Inmetro`
   - Corpo do email: fará uma cópia das informações dos bolsistas:
   <img width="400" height="500" alt="image" src="https://github.com/user-attachments/assets/971aee8d-8f76-46c9-bc06-4f2c714bb2b8" />

8. Depois pode digitar `s` para enviar.
