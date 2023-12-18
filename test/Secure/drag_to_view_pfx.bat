@openssl pkcs12 -in %1 -nodes -passin pass:"" | openssl x509 -text
@pause
