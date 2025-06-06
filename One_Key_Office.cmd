::    One_Click_Office: Install & activate Microsoft Office with one click
::    Copyright (C) 2024  Charon2050
::
::    This program is free software: you can redistribute it and/or modify
::    it under the terms of the GNU General Public License as published by
::    the Free Software Foundation, either version 3 of the License, or
::    (at your option) any later version.
::
::    This program is distributed in the hope that it will be useful,
::    but WITHOUT ANY WARRANTY; without even the implied warranty of
::    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
::    GNU General Public License for more details.
::
::    You should have received a copy of the GNU General Public License
::    along with this program.  If not, see <https://www.gnu.org/licenses/>.

@echo off
chcp 65001

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 未以管理员权限运行，正在请求管理员权限...
    PowerShell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    exit /b
)

cd /D %temp%

echo ^<Configuration ID="90ccfe8f-4187-43db-b0cc-aea40cd487fe"^>^<Add OfficeClientEdition="64" Channel="Current"^>^<Product ID="O365ProPlusRetail"^>^<Language ID="MatchOS" /^>^<ExcludeApp ID="Access" /^>^<ExcludeApp ID="Groove" /^>^<ExcludeApp ID="Lync" /^>^<ExcludeApp ID="OneDrive" /^>^<ExcludeApp ID="OneNote" /^>^<ExcludeApp ID="Outlook" /^>^<ExcludeApp ID="Publisher" /^>^<ExcludeApp ID="Teams" /^>^<ExcludeApp ID="Bing" /^>^</Product^>^</Add^>^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>^<Updates Enabled="TRUE" /^>^<RemoveMSI /^>^</Configuration^> > Configuration.xml


@if not exist "OfficeDeployTool.exe" (
    echo 正在从 officecdn.microsoft.com 下载微软官方 Office 安装器...
    curl https://officecdn.microsoft.com/pr/wsus/setup.exe --output OfficeDeployTool.exe
)

@if not exist "OfficeDeployTool.exe" (
    echo 下载失败，请检查网络，并关闭杀毒软件，然后重启本程序。
    pause
    exit /b
)

echo 开始安装 Office...
OfficeDeployTool /configure Configuration.xml

echo 正在激活...
powershell -Command "& {& ([ScriptBlock]::Create((irm https://get.activated.win))) /Ohook}"

echo,
echo,
echo Office 安装并激活完成！
pause
