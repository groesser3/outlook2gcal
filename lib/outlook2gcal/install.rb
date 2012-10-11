#encoding: utf-8
require 'rubygems'
require 'log4r'
require 'fileutils'

module Outlook2GCal
  
  def Outlook2GCal::install(conf)
    raise "environment variable %APPDATA% not set" unless ENV['APPDATA'] 
    puts File.realdirpath(__FILE__+ '/../..')
    puts File.realdirpath($0+ '/../..')
    puts "---"
    
    src_dir = File.realdirpath($0 + '/../..') + '/.'
    dest_dir = File.realdirpath(ENV['APPDATA']+"/outlook2gcal_standalone_0.0.1")
    if !File.exists?(dest_dir)
      copy_files(src_dir,dest_dir)
      create_licence_file(dest_dir)
      create_start_file(dest_dir)
    else
      puts "Dir (#{dest_dir}) already exists. "
    end  

  end
  
  def Outlook2GCal::copy_files(src_dir, dest_dir)
    puts "Copying files from " + src_dir
    puts "to " + dest_dir
    FileUtils::mkdir_p(dest_dir)
    FileUtils::cp_r(src_dir, dest_dir)
    puts "Install successful!" 
  end
  def Outlook2GCal::create_licence_file(dest_dir)
    licence =<<__END__
Copyright (c) 2012 Elias Kugler

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
__END__

    f= File.new(dest_dir+"/MIT-LICENSE",'w+') 
    f.write(licence)
    f.close
  end 
  
  def Outlook2GCal::create_start_file(dest_dir)
    cmd =<<__END__
@goto :end                ;@rem TODO: delete this line, edit the next one !!!
.\\bin\\ruby .\\src\\outlook2gcal_portable sync -U <google-user> -P <google-password> -C <calendar-name> -D 30  --sync-desc --sync-alarm
@echo.
@pause
@rem TODO: delete the following lines !!!

:end                          
@echo off                                                                     
echo --------------------------
echo  TODO!!! Edit this file!!                                             
echo --------------------------
.\\bin\\ruby .\\src\\outlook2gcal_portable   
echo --------------------------
echo  TODO!!! Edit this file!!                                             
echo --------------------------
@pause
__END__

    f= File.new(dest_dir+"/start_outlook2gcal_portable.bat",'w+') 
    f.write(cmd)
    f.close
  end 
end

    if __FILE__ == $0 then 
      puts "test"
      Outlook2GCal::create_start_file(File.realdirpath('.'))
    end