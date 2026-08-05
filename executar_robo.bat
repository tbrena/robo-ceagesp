@echo off
REM Executa o robo de cotacoes da CEAGESP (usado pelo Agendador de Tarefas do Windows)
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

set PY=python
where py >nul 2>&1 && set PY=py

where %PY% >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Python nao encontrado no PATH.
  echo Instale o Python 3 ou edite a variavel PY neste arquivo com o caminho completo do python.exe.
  exit /b 9009
)

%PY% "%~dp0robo_ceagesp.py" %*
exit /b %ERRORLEVEL%
