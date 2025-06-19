gulp trust-dev-cert

cp ~/.rushstack/rushstack-serve.key ./spfx-dev-cert.key
cp ~/.rushstack/rushstack-serve.pem ./spfx-dev-cert.pem

## run outside of devcontainer
certutil -user -addstore root "C:\Rachael\projects\react-directory\spfx-dev-cert.pem"
