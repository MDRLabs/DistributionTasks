@echo Off
npm config set proxy http://nl-userproxy-access.net.abnamro.com:8080
npm config set http-proxy http://nl-userproxy-access.net.abnamro.com:8080
npm config set https-proxy http://nl-userproxy-access.net.abnamro.com:8080
npm config set https_proxy http://nl-userproxy-access.net.abnamro.com:8080
npm config get proxy
npm config get http-proxy
npm config get https-proxy
npm config get http_proxy
npm config get https_proxy
@echo On
