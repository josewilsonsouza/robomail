# Automatizar o Envio Semanal

Este guia explica como configurar o Windows Task Scheduler para executar o script automaticamente toda quinta-feira às 14:30.

## Task Scheduler

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

Para testar se a tarefa funciona:

1. No Task Scheduler, encontre a tarefa criada
2. Clique com botão direito → "Executar"
3. Verifique se o script abre e funciona corretamente

**IMPORTANTE**: O Task Scheduler executará o script em segundo plano. Para ver a janela durante a execução automática, você precisa estar logado no Windows.

---

### Erro "Python não encontrado"
- Edite o arquivo `.bat` para usar o caminho completo do Python:
```batch
@echo off
cd /d "%~dp0"
"C:\Python39\python.exe" mail.py
pause
```

# INSTRUÇÕES PARA LINUX

Se você estiver usando Linux, siga estas adaptações:

## Ajuste no Código (mail.py)

Substitua a linha 378 por uma versão multiplataforma:

```python
#(Windows específico):
file_path_url = f"file:///{html_temp_path.replace(chr(92), '/')}"

# Multiplataforma:
import platform
if platform.system() == "Windows":
    file_path_url = f"file:///{html_temp_path.replace(chr(92), '/')}"
else:
    file_path_url = f"file://{html_temp_path}"
```

## Automatização com Cron (Linux)

### Criar Script Shell

Crie um arquivo `executar_envio_email.sh`:

```bash
#!/bin/bash
cd "$(dirname "$0")"
python3 mail.py
```

### Dar permissão de execução

```bash
chmod +x executar_envio_email.sh
```

### Configurar Cron

Abra o editor de cron:
```bash
crontab -e
```

Adicione esta linha (executa toda quinta-feira às 14:30):
```
30 14 * * 4 /caminho/completo/para/executar_envio_email.sh
```

### Verificar se o cron está ativo

```bash
# Ver tarefas agendadas
crontab -l

# Ver status do serviço cron
sudo systemctl status cron
```

### Teste

```bash
./executar_envio_email.sh
```

O `webdriver-manager` detecta automaticamente qual está instalado.

**Instalar Chrome no Ubuntu/Debian:**
```bash
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo dpkg -i google-chrome-stable_current_amd64.deb
sudo apt-get install -f
```

**usar Chromium:**
```bash
sudo apt install chromium-browser
```