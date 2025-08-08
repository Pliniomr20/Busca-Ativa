@echo off
title Dashboard Busca Ativa

REM Altera para o diretório onde o script e a planilha estão
cd /d C:\Users\USUARIO\Desktop\BUSCA ATIVA

REM Inicia o dashboard do Streamlit
streamlit run dashboard_busca_ativa.py

REM Pausa o script para que a janela não feche imediatamente após o erro
pause