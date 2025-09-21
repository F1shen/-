@echo off
setlocal
cd /d %~dp0

start "" python\python-3.13.7-embed-amd64\python.exe -m streamlit run app.py --server.port 8501 --server.headless true

rem ==== 固定走 8501，關掉 dev-mode（避免 :3000） ====
set STREAMLIT_GLOBAL_DEVELOPMENTMODE=false
set STREAMLIT_BROWSER_GATHERUSAGESTATS=false
set STREAMLIT_SERVER_HEADLESS=true
set STREAMLIT_SERVER_PORT=8501

rem ==== 後台啟動 Streamlit 伺服器 ====
start "" python -m streamlit run app.py --server.port 8501 --server.headless true

rem ==== 用 PowerShell 等 8501 通了，再打開瀏覽器 ====
powershell -NoProfile -Command ^
  "$p=8501; $ok=$false; for($i=0;$i -lt 60;$i++){try{(New-Object Net.Sockets.TcpClient('127.0.0.1',$p)).Close(); $ok=$true; break}catch{}; Start-Sleep -Milliseconds 500}; if($ok){ Start-Process 'http://localhost:8501' }"

endlocal
