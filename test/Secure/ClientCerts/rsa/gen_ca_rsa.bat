@echo off

openssl genrsa -out ca_rsa.key 4096
openssl req -new -x509 -days 3650 -key ca_rsa.key -out ca_rsa.pem
