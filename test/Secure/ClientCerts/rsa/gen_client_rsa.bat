@echo off
setlocal
set "CLIENT_ID=%~1"
set "CLIENT_SERIAL=%~2"

if [%CLIENT_ID%]==[] echo usage: %~nx0 ^<CLIENT_ID^> ^<CLIENT_SERIAL^>& exit /b 1

openssl genrsa -out "%CLIENT_ID%.key" 2048
openssl req -new -key "%CLIENT_ID%.key" -out "%CLIENT_ID%.csr"
openssl x509 -req -days 3650 -in "%CLIENT_ID%.csr" -CA ca_rsa.pem -CAkey ca_rsa.key -set_serial %CLIENT_SERIAL% -out "%CLIENT_ID%.pem"
openssl pkcs12 -export -out %CLIENT_ID%.full.pfx -inkey "%CLIENT_ID%.key" -in "%CLIENT_ID%.pem" -certfile ca_rsa.pem
