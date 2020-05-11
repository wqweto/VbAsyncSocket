@echo off

openssl ecparam -genkey -name secp256r1 | openssl ec -out ca.key
openssl req -new -x509 -days 3650 -key ca.key -out ca.pem
