@echo off
echo Excel To JSON/YAML Converter Add-in ?ㅼ튂 以?..

REM ?ㅼ튂 ?붾젆?좊━ ?앹꽦
set INSTALL_DIR=%APPDATA%\Microsoft\AddIns\ExcelToYamlAddin
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM ?뚯씪 蹂듭궗
xcopy /Y /E "%~dp0Release\*.*" "%INSTALL_DIR%\"

echo ?ㅼ튂媛 ?꾨즺?섏뿀?듬땲??
echo Excel???닿퀬 ?뚯씪 > ?듭뀡 > 異붽? 湲곕뒫?쇰줈 ?대룞?섏뿬 ?좊뱶?몄씠 ?쒖꽦?붾릺?덈뒗吏 ?뺤씤?섏꽭??
pause
