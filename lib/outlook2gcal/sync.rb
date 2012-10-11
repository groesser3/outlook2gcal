#encoding: utf-8
require 'rubygems'
require 'log4r'
require 'uri'
require 'date'
require 'fileutils'
require 'outlook2gcal/outlook_calendar'
require 'outlook2gcal/google_calendar'
require 'outlook2gcal/string'
require 'outlook2gcal/counter'
require 'outlook2gcal/sync_entry'


module Outlook2GCal
  
  class OutlookGoogleSync
    def initialize(params)
      @params = params
      @google_calendar = Outlook2GCal::GoogleCalendar.new(params)
      
      @sync_time = nil
      if params[:days]
         @min_sync_time = DateTime.now-params[:days] #*86400
      end
      if params[:days_max]
         @max_sync_time = DateTime.now+params[:days_max] #*86400
      else 
         @max_sync_time = DateTime.now+400 #*86400
      end
      @max_time = DateTime.parse("2038-01-18")
      # do not sync the description unless the users wants to
      @sync_desc = params[:sync_desc] # || true
      @sync_alarm = params[:sync_alarm]
      @sync_names = params[:sync_names]
      
      init_logger
    end
    def init_logger
      FileUtils::mkdir_p('Outlook2GCal')
      $logger = Log4r::Logger.new("sync_logger")      
      Log4r::FileOutputter.new('logfile', 
                         :filename=>"#{Dir.pwd}/Outlook2GCal/Outlook2GCal.log", 
                         :trunc=>false,
                         :level=>Log4r::WARN)
      $logger.add('logfile')
    end
    def sync_events
      @counter = Counter.new
      @sync_time = DateTime.now
      sleep(0)
      @outlook_calendar = Outlook2GCal::OutlookCalendar.new(@params)
      @outlook_calendar.events.each do |outlook_event|
        @counter.selects += 1
   #     if outlook_event.repeats?
   #       outlook_event.repeats.each do |r|
   #         sync_event(outlook_event, r.start_time, r.end_time, outlook_event.uid+'_'+r.start_time)
   #       end
   #     else
          #p outlook_event.repeats
          sync_event(outlook_event, outlook_event.start_time, outlook_event.end_time, outlook_event.uid)
    #    end
      end
      @outlook_calendar.quit
      del_events()
      @counter.end
      return @counter
    end

    def sync_event(outlook_event, start_time, end_time, key)
        sdt = DateTime.parse(start_time)
        return unless sdt < @max_time # workaround 

        if end_time
          edt = DateTime.parse(end_time)
          return unless edt < @max_time # workaround 
        end
        
        #puts DateTime.parse(outlook_event.end_time)
        
        if (@min_sync_time and end_time and 
           @min_sync_time > DateTime.parse(end_time)) or
          (@max_sync_time and start_time and
            @max_sync_time < DateTime.parse(start_time))
        then
          @counter.ignored +=1
        else
          #p key
          sync_entry = Outlook2GCal::SyncEntry.first(:outlook_uid => key) 
          if sync_entry  
          then
            #puts DateTime.parse(outlook_event.end_time)
            #p sync_entry.outlook_last_modified.to_s
            #p outlook_event.last_modified
            if sync_entry.outlook_last_modified < outlook_event.last_modified
            then
              #!!insert_update(sync_entry,outlook_event)
              insert_update(sync_entry,outlook_event, start_time, end_time, key)
            else
              print "."
              @counter.ignored +=1
              sync_entry.sync_time = @sync_time
              sync_entry.sync_action = 'N' # none
              sync_entry.save
            end
          else
            add_event(outlook_event, start_time, end_time, key)
          end
        end
    end

    def del_events
        Outlook2GCal::SyncEntry.all(:sync_time.lt => @sync_time).each do |sync_entry|
          @counter.deletes += 1
          if @google_calendar.del_event(sync_entry.gcal_id)
            print "D"
            sync_entry.destroy
          else 
            sync_entry.sync_time = @sync_time
            sync_entry.sync_action = 'E'
            sync_entry.save
            print "E"
          end
        end
    end
    def init_google_event(outlook_event,start_time,end_time)
      event = @google_calendar.new_event
      google_event= set_google_event_attrs(outlook_event, event,start_time,end_time )
      google_event.start_time = start_time 
      google_event.end_time = end_time 
      return google_event
    end
    def set_google_event_attrs(outlook_event, google_event,start_time=nil,end_time=nil)
      google_event.title = outlook_event.subject.asciify if outlook_event.subject
      if start_time
        google_event.start_time = start_time     
      else
        google_event.start_time = outlook_event.start_time 
      end
      if end_time
        google_event.end_time = end_time 
      else
        google_event.end_time = outlook_event.end_time 
      end
      google_event.where = outlook_event.where.asciify if outlook_event.where
      google_event.all_day = outlook_event.all_day?
      
      if (@sync_desc || @sync_names)
        content = ''
        content += outlook_event.formatted_names.asciify if @sync_names
        content += outlook_event.content.asciify if @sync_desc      
        google_event.content = content 
        #puts content
      end
      
      if @sync_alarm and outlook_event.alarm
        google_event.reminder = [{:method =>'alert', :minutes => outlook_event.alarm_offset }]
      end
      
      return google_event
    end
    def get_sync_entry_by_notes_uid(uid)
      e1 = Outlook2GCal::SyncEntry.first(:outlook_uid => uid)
      return e1
    end
    def insert_update(sync_entry,outlook_event, start_time, end_time, key)
      gcal_event = @google_calendar.find_event(sync_entry.gcal_id)
      if gcal_event == []
        $logger.warn "Event not found for update"
        add_event(outlook_event,start_time, end_time,key)
      else 
        update_event(sync_entry, outlook_event, gcal_event, start_time, end_time)
      end
    end
   
    def add_event(outlook_event, start_time, end_time, key)
      print "I"
      google_event=init_google_event(outlook_event, start_time, end_time)
      #p google_event
      ret = google_event.save
      $logger.fatal "insert: cannot save gcal event" unless ret
      raise "cannot save gcal event" unless ret
      @counter.inserts +=1
      sync_entry = Outlook2GCal::SyncEntry.new
      sync_entry.outlook_uid = key #outlook_event.uid
      sync_entry.sync_time = @sync_time
      sync_entry.outlook_last_modified = outlook_event.last_modified
      sync_entry.gcal_id = google_event.id
      sync_entry.sync_action = 'I' # insert
      sync_entry.save
    end
    def update_event(sync_entry, outlook_event, gcal_event, start_time, end_time)
      print "U"
      @counter.updates +=1
      set_google_event_attrs(outlook_event, gcal_event, start_time, end_time)
      ret = gcal_event.save
      $logger.fatal "update: cannot save gcal event" unless ret
      raise "cannot save gcal event" unless ret
      sync_entry.sync_time = @sync_time
      sync_entry.gcal_id = gcal_event.id
      sync_entry.outlook_last_modified = outlook_event.last_modified
      sync_entry.sync_action = 'U' # none
      sync_entry.save
    end
  end
  
end