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

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 未以管理员权限运行，正在请求管理员权限...
    PowerShell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    exit /b
)

cd /D %temp%

echo ^<Configuration ID="90ccfe8f-4187-43db-b0cc-aea40cd487fe"^>^<Add OfficeClientEdition="64" Channel="Current"^>^<Product ID="O365ProPlusRetail"^>^<Language ID="MatchOS" /^>^<ExcludeApp ID="Access" /^>^<ExcludeApp ID="Groove" /^>^<ExcludeApp ID="Lync" /^>^<ExcludeApp ID="OneDrive" /^>^<ExcludeApp ID="OneNote" /^>^<ExcludeApp ID="Outlook" /^>^<ExcludeApp ID="Publisher" /^>^<ExcludeApp ID="Teams" /^>^<ExcludeApp ID="Bing" /^>^</Product^>^</Add^>^<Property Name="SharedComputerLicensing" Value="0" /^>^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>^<Property Name="DeviceBasedLicensing" Value="0" /^>^<Property Name="SCLCacheOverride" Value="0" /^>^<Updates Enabled="TRUE" /^>^<RemoveMSI /^>^</Configuration^> > Configuration.xml


@if not exist "setup.exe" (
    echo 正在从 officecdn.microsoft.com 下载微软官方 Office 安装器...
    certutil -urlcache -split -f https://officecdn.microsoft.com/pr/wsus/setup.exe
)

echo 开始安装 Office...
setup.exe /configure Configuration.xml

@if not exist "Ohook_Activation_AIO.cmd" (
    echo 正在从 dev.azure.com 下载 MAS 激活器...
    certutil -urlcache -split -f https://dev.azure.com/massgrave/e5e7ca26-b96c-4887-b3d8-ce10a5bdee05/_apis/git/repositories/a5c49950-3666-490e-afa6-1c2c2d29815e/items?path=/MAS/Separate-Files-Version/Activators/Ohook_Activation_AIO.cmd&versionDescriptor%5BversionOptions%5D=0&versionDescriptor%5BversionType%5D=0&versionDescriptor%5Bversion%5D=master&resolveLfs=true&%24format=octetStream&api-version=5.0&download=true
)

echo 正在激活...
call Ohook_Activation_AIO.cmd /Ohook

echo,
echo,
echo Office 安装并激活完成！
pause