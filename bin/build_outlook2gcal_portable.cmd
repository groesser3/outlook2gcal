rem build ncal2gcal portable using the ocra gem
rem installing ocra: gem install ocra
cp outlook2gcal "%appdata%\outlook2gcal_portable"
cd /d %appdata%
ocra outlook2gcal_portable
pause
