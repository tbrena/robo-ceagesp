@echo off
REM Executa o robo de cotacoes da CEAGESP (usado pelo Agendador de Tarefas do Windows)
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
"C:\Users\thiago.IEA\AppData\Local\Programs\Python\Python312\python.exe" "%~dp0robo_ceagesp.py" %*
exit /b %ERRORLEVEL%
