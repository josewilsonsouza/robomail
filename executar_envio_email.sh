#!/bin/bash
# Script para executar o envio automático de email de transporte
# Funciona em Linux/Mac

# Muda para o diretório do script
cd "$(dirname "$0")"

# Executa o script Python
python3 mail.py

# Pausa para ver resultado (opcional, comente se rodar via cron)
read -p "Pressione ENTER para fechar..."
