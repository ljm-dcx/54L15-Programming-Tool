@ECHO OFF
set TOOL_CHAIN_DIR=C:\ncs\toolchains\0b393f9e1b
set BASE=C:\ncs
set ZEPHYR_BASE=%BASE%\v3.0.2\zephyr
set ZEPHYR_SDK_INSTALL_DIR=%TOOL_CHAIN_DIR%\opt\zephyr-sdk
set PYTHONPATH=%TOOL_CHAIN_DIR%\opt\bin;%TOOL_CHAIN_DIR%\opt\bin\Lib;%TOOL_CHAIN_DIR%\opt\bin\Lib\site-packages
set PATH=%TOOL_CHAIN_DIR%;%TOOL_CHAIN_DIR%\mingw64\bin;%TOOL_CHAIN_DIR%\bin;%TOOL_CHAIN_DIR%\opt\bin;%TOOL_CHAIN_DIR%\opt\bin\Scripts;%TOOL_CHAIN_DIR%\opt\nanopb\generator-bin;%TOOL_CHAIN_DIR%\opt\zephyr-sdk\arm-zephyr-eabi\bin;%PATH%
@ECHO ------------------------------------------------
@ECHO   NCS tool chain v3.0.2
@ECHO   NCS %BASE%
REM @ECHO   NCS %PATH%
@ECHO ------------------------------------------------
python --version
REM cd %BASE%
REM nrfutil device erase --all
west ncs-provision upload -k configuration/nrf54l15dk_nrf54l15_cpuapp/boot_signature_key_file_ed25519.pem --keyname UROT_PUBKEY -s nrf54l15
