cd /d D:\home\elkugler\devel\Sonstiges\ruby\outlook2gcal
gem build outlook2gcal.gemspec
gem install outlook2gcal-0.0.1.gem --local

cd /d D:\home\elkugler\devel\Sonstiges\ruby\outlook2gcal\bin
call build_outlook2gcal_portable.cmd

cd D:\Users\elkugler\AppData\Roaming
outlook2gcal_portable install