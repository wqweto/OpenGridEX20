@echo off
reg.exe delete "HKEY_CLASSES_ROOT\TypeLib\{24F4AB9F-37F4-43D4-B383-FB6CD721B629}" /f
regsvr32 /s %~dp0\OpenGridEX20.ocx
