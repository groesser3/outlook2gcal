#encoding: utf-8
require 'dm-core'
require 'dm-migrations'
require 'fileutils'

module Outlook2GCal
  raise "environment variable %APPDATA% not set" unless ENV['APPDATA'] 
  FileUtils::cd(ENV['APPDATA']) 
  FileUtils::mkdir_p('outlook2gcal')
  
  db_file = "#{Dir.pwd}/outlook2gcal/outlook2gcal.sqlite"
  DataMapper::setup(:default, "sqlite3://#{db_file}")

  
  class SyncEntry
    include DataMapper::Resource
    #storage_names[:repo] = 'ncal2gal_sync_entries'

    property :id, Serial
    property :sync_time, DateTime
    property :sync_action, String
    
    property :outlook_uid, Text, :index=>true
    property :outlook_last_modified, DateTime  
    property :gcal_id, Text

  end
  
  # automatically create the SyncEntry table
  SyncEntry.auto_migrate! unless File.exists?(db_file) #SyncEntry.table_exists?
end