ARG WINDOWS_VER

FROM mcr.microsoft.com/dotnet/framework/runtime:4.8-windowsservercore-ltsc2019
SHELL ["powershell", "-command"]

COPY ./testZip.zip .
COPY ./test.vbs .
