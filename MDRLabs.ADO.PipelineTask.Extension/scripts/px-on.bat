@echo Off
npm config set proxy your-proxy-goes-here
npm config set http-proxy your-proxy-goes-here
npm config set https-proxy your-proxy-goes-here
npm config set https_proxy your-proxy-goes-here
npm config get proxy
npm config get http-proxy
npm config get https-proxy
npm config get http_proxy
npm config get https_proxy
@echo On
