@echo off
npm config set http_proxy 
npm config set https_proxy 
npm config get http_proxy
npm config get https_proxy
