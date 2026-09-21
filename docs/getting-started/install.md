# Установка и сборка

## 1. Сборка DLL

### Вариант A — Visual Studio

1. Откройте `AcSapfir.sln`.
2. Выберите конфигурацию `Debug` или `Release` (платформа `AnyCPU`).
3. Соберите решение (**Build → Build Solution**, `F6`).

### Вариант B — MSBuild из командной строки

```powershell
msbuild AcSapfir.csproj /t:Build /p:Configuration=Debug
```

!!! warning "Используйте MSBuild из Visual Studio, а не `dotnet build`"
    Проект содержит **COM-ссылку** на `SapfirLib` (`<COMReference>` с `WrapperTool=tlbimp`). Классический `msbuild` (`C:\Program Files\Microsoft Visual Studio\2022\...\MSBuild\Current\Bin\MSBuild.exe`) её обрабатывает, а `dotnet build` — нет.

Результат сборки:

```
bin\Debug\AcSapfir.dll      (или bin\Release\AcSapfir.dll)
```

### Если сборка падает на ссылках AutoCAD

В `AcSapfir.csproj` пути к библиотекам AutoCAD заданы абсолютно:

```xml
<Reference Include="accoremgd">
  <HintPath>..\..\..\..\..\Program Files\Autodesk\AutoCAD 2020\accoremgd.dll</HintPath>
  <Private>False</Private>
</Reference>
```

Если AutoCAD установлен в другую папку (или это другая версия), поправьте `HintPath` для `accoremgd`, `acdbmgd`, `acmgd`. Флаг `Private=False` означает, что API-библиотеки AutoCAD **не** копируются в `bin` — они берутся из установленного AutoCAD в момент выполнения.

### Если сборка падает на COM-ссылке

```xml
<COMReference Include="SapfirLib">
  <Guid>{94863C7A-F77E-418F-9F50-3C0AD68A509A}</Guid>
  <VersionMajor>19</VersionMajor>
  <VersionMinor>0</VersionMinor>
</COMReference>
```

Требуется установленный САПФИР, зарегистрировавший эту тип-библиотеку. Проверка — в разделе [Требования](requirements.md).

## 2. Загрузка в AutoCAD

1. Запустите AutoCAD 2020 и САПФИР (САПФИР — обязательно, до вызова команд плагина).
2. В AutoCAD выполните `NETLOAD` и укажите `bin\Debug\AcSapfir.dll`.
3. Проверьте доступность команд: наберите, например, `Sapfir_AXES` — AutoCAD должен распознать команду.

Регистрация команд происходит автоматически за счёт атрибута в сборке:

```csharp
[assembly: CommandClass(typeof(AcSapfir.SapfirDrafter))]
```

При повторной сборке DLL просто загрузите её заново командой `NETLOAD` (AutoCAD удерживает старую сборку в памяти — если команда не обновляется, перезапустите AutoCAD).

## 3. Автозагрузка при старте AutoCAD (необязательно)

Добавьте запись в реестр, чтобы плагин загружался автоматически:

```powershell
$key = 'HKCU:\Software\Autodesk\AutoCAD\R23.1\ACAD-3101:409\Applications\AcSapfir'
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name DESCRIPTION -Value 'AcSapfir'            -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name LOADER      -Value 'C:\path\bin\Debug\AcSapfir.dll' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name LOADCTRLS   -Value 2 -PropertyType DWord -Force | Out-Null
```

| Значение | Смысл |
|---|---|
| `R23.1` | ветка реестра AutoCAD 2020 |
| `ACAD-3101:409` | код продукта и язык; проверьте список в `HKCU\Software\Autodesk\AutoCAD\R23.1\` |
| `LOADCTRLS = 2` | загружать плагин при старте AutoCAD |

После правки реестра перезапустите AutoCAD.

## 4. Диагностика загрузки

| Симптом | Причина / что делать |
|---|---|
| `NETLOAD` не открывает диалог | Проверьте `FILEDIA` (=1), введите `NETLOAD` в командной строке |
| «Не удаётся загрузить сборку» | Соберите под .NET Framework 4.8 и убедитесь, что DLL не заблокирована (`Unblock-File .\AcSapfir.dll`) |
| Команда не найдена после загрузки | DLL загружена из другого места; перезапустите AutoCAD |
| При вызове команды исключение COM | САПФИР не запущен либо COM-библиотека не зарегистрирована — см. [Решение проблем](../guides/troubleshooting.md) |
