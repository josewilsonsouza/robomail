# Como Automatizar o Envio Semanal

Este guia explica como configurar o Windows Task Scheduler para executar o script automaticamente toda quinta-feira às 14:30.

## Opção 1: Task Scheduler (Recomendado)

### Passo 1: Criar Arquivo Batch

Primeiro, crie um arquivo `.bat` para executar o script Python:

1. Crie um novo arquivo chamado `executar_envio_email.bat` na mesma pasta do `mail.py`
2. Adicione o seguinte conteúdo:

```batch
@echo off
cd /d "%~dp0"
python mail.py
pause
```

3. Salve o arquivo

### Passo 2: Configurar Task Scheduler

1. **Abrir Task Scheduler**:
   - Pressione `Win + R`
   - Digite `taskschd.msc`
   - Pressione Enter

2. **Criar Nova Tarefa**:
   - Clique em "Criar Tarefa..." (no painel direito)
   - NÃO use "Criar Tarefa Básica"

3. **Aba "Geral"**:
   - Nome: `Envio Email Transporte Inmetro`
   - Descrição: `Envia email semanal de solicitação de transporte toda quinta às 14:30`
   - Marque: "Executar estando o usuário conectado ou não"
   - Marque: "Executar com privilégios mais altos" (se necessário)
   - Configure para: `Windows 10`

4. **Aba "Disparadores"**:
   - Clique em "Novo..."
   - Configurações:
     - Iniciar a tarefa: `Seguindo um agendamento`
     - Configurações: `Semanalmente`
     - Recorrência: `Repetir a cada: 1 semana`
     - Dias da semana: Marque apenas `Quinta-feira`
     - Horário: `14:30:00`
     - Habilitado: `✓`
   - Clique em "OK"

5. **Aba "Ações"**:
   - Clique em "Novo..."
   - Ação: `Iniciar um programa`
   - Programa/script: Clique em "Procurar..." e selecione o arquivo `executar_envio_email.bat`
   - Iniciar em: Deixe em branco (será a pasta do .bat)
   - Clique em "OK"

6. **Aba "Condições"**:
   - DESMARQUE: "Iniciar a tarefa apenas se o computador estiver conectado à energia CA"
   - DESMARQUE: "Parar se o computador alternar para energia de bateria"
   - Marque: "Ativar o computador para executar esta tarefa" (se quiser que o PC acorde)

7. **Aba "Configurações"**:
   - Marque: "Permitir que a tarefa seja executada sob demanda"
   - Marque: "Executar a tarefa assim que possível após uma inicialização agendada ser perdida"
   - Se a tarefa falhar, reiniciar a cada: `1 minuto`
   - Tentar reiniciar até: `3 vezes`

8. **Finalizar**:
   - Clique em "OK"
   - Se solicitado, digite a senha do usuário Windows

### Passo 3: Testar a Tarefa

Para testar se a tarefa funciona:

1. No Task Scheduler, encontre a tarefa criada
2. Clique com botão direito → "Executar"
3. Verifique se o script abre e funciona corretamente

**IMPORTANTE**: O Task Scheduler executará o script em segundo plano. Para ver a janela durante a execução automática, você precisa estar logado no Windows.

---

## Opção 2: Script Python com Loop (Alternativa)

Se preferir não usar Task Scheduler, pode criar um script que fica sempre rodando:

Crie um arquivo `monitor_semanal.py`:

```python
import time
from datetime import datetime
import subprocess

def executar_se_quinta_1430():
    """Executa o script mail.py se for quinta às 14:30"""
    agora = datetime.now()

    # Quinta-feira = 3 (0=segunda, 1=terça, etc.)
    if agora.weekday() == 3:  # Quinta
        if agora.hour == 14 and agora.minute == 30:
            print(f"🚀 {agora.strftime('%d/%m/%Y %H:%M')} - Executando envio de email...")
            try:
                subprocess.run(["python", "mail.py"])
                print("✅ Execução concluída!")
            except Exception as e:
                print(f"❌ Erro: {e}")

            # Aguarda 2 minutos para não executar múltiplas vezes
            time.sleep(120)

print("🔄 Monitor iniciado. Aguardando quinta-feira às 14:30...")
print("Pressione Ctrl+C para parar.")

while True:
    executar_se_quinta_1430()
    time.sleep(30)  # Verifica a cada 30 segundos
```

**Desvantagens**:
- Precisa deixar o script rodando 24/7
- Consome recursos constantemente
- Não é tão confiável quanto Task Scheduler

---

## Opção 3: Execução Manual com Lembrete

Se preferir controle manual:

1. Configure um lembrete no Outlook/Google Calendar para toda quinta às 14:25
2. Execute manualmente: `python mail.py` toda quinta-feira

**Vantagem**: Você sempre valida antes de enviar
**Desvantagem**: Depende de lembrar

---

## Solução de Problemas

### Tarefa não executa no horário
- Verifique se o computador está ligado às 14:30 (quinta)
- Verifique se o usuário Windows está logado
- No Task Scheduler: Clique com direito na tarefa → "Histórico" para ver logs

### Erro "Python não encontrado"
- Edite o arquivo `.bat` para usar o caminho completo do Python:
```batch
@echo off
cd /d "%~dp0"
"C:\Python39\python.exe" mail.py
pause
```

### Credenciais expiram
- Se suas credenciais de rede mudam periodicamente, você precisará executar manualmente quando isso acontecer
- O script sempre pede usuário/senha interativamente

### Email não envia automaticamente
- O script pede confirmação antes de enviar
- Para automação total, você precisaria modificar o código para pular confirmações (NÃO recomendado)

---

## Recomendação Final

**Use a Opção 1 (Task Scheduler)** porque:
- ✅ Nativo do Windows, confiável
- ✅ Não consome recursos quando não está rodando
- ✅ Histórico e logs de execução
- ✅ Pode executar mesmo com usuário deslogado
- ✅ Configuração única, funciona para sempre

---

# INSTRUÇÕES PARA LINUX

Se você estiver usando Linux, siga estas adaptações:

## Ajuste no Código (mail.py)

Substitua a linha 378 por uma versão multiplataforma:

```python
# ANTES (Windows específico):
file_path_url = f"file:///{html_temp_path.replace(chr(92), '/')}"

# DEPOIS (Multiplataforma):
import platform
if platform.system() == "Windows":
    file_path_url = f"file:///{html_temp_path.replace(chr(92), '/')}"
else:
    file_path_url = f"file://{html_temp_path}"
```

## Automatização com Cron (Linux)

### Passo 1: Criar Script Shell

Crie um arquivo `executar_envio_email.sh`:

```bash
#!/bin/bash
cd "$(dirname "$0")"
python3 mail.py
```

### Passo 2: Dar permissão de execução

```bash
chmod +x executar_envio_email.sh
```

### Passo 3: Configurar Cron

Abra o editor de cron:
```bash
crontab -e
```

Adicione esta linha (executa toda quinta-feira às 14:30):
```
30 14 * * 4 /caminho/completo/para/executar_envio_email.sh
```

**Explicação do formato cron:**
```
30 14 * * 4
│  │  │ │ └─── Dia da semana (0=domingo, 4=quinta)
│  │  │ └───── Mês (1-12)
│  │  └─────── Dia do mês (1-31)
│  └────────── Hora (14 = 14h)
└───────────── Minuto (30)
```

### Passo 4: Verificar se o cron está ativo

```bash
# Ver tarefas agendadas
crontab -l

# Ver status do serviço cron
sudo systemctl status cron
```

### Passo 5: Testar manualmente

```bash
./executar_envio_email.sh
```

## Diferenças Chrome/Chromium no Linux

No Linux, você pode usar:
- **Google Chrome** (mesmo do Windows)
- **Chromium** (versão open source)

O `webdriver-manager` detecta automaticamente qual está instalado.

**Instalar Chrome no Ubuntu/Debian:**
```bash
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo dpkg -i google-chrome-stable_current_amd64.deb
sudo apt-get install -f
```

**Ou usar Chromium:**
```bash
sudo apt install chromium-browser
```

## Problemas Comuns no Linux

### 1. Display não encontrado (headless mode)

Se rodar via cron sem interface gráfica, pode dar erro. Solução:

Adicione no código (após linha 193):
```python
options.add_argument("--headless")  # Roda sem abrir janela
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
```

**PORÉM**: Modo headless pode ter problemas com a janela popup do email. Ideal é rodar com interface gráfica.

### 2. Permissões de arquivo

```bash
chmod 644 mail.py
chmod 755 executar_envio_email.sh
```

### 3. Variável DISPLAY para cron

Se precisar rodar com interface gráfica via cron:
```
30 14 * * 4 DISPLAY=:0 /caminho/completo/para/executar_envio_email.sh
```

## Logs de Execução (Cron)

Para salvar logs:
```
30 14 * * 4 /caminho/completo/para/executar_envio_email.sh >> /var/log/email_transporte.log 2>&1
```
