#encoding: utf-8
require 'win32ole'

# Info: 
#     http://msdn.microsoft.com/en-us/library/dd469461(v=office.12).aspx
#     http://msdn.microsoft.com/en-us/library/ff870661.aspx

module Outlook2GCal
  class EventRepeat
    attr_accessor   :start_time, :end_time
    def initialize
    end
  end
  class OutlookCalendar
    attr_accessor :server, :user, :password, :db 
    attr_reader :events
    def initialize(params)
      @server = params[:notes_server] || '' # local
      @user = params[:notes_user] 
      @password = params[:notes_password] 
      @db = params[:notes_db] 
      
      created = false
      @ol = nil
      begin 
        @ol = WIN32OLE.connect("Outlook.Application") 
      rescue 
        @created = true
        @ol = WIN32OLE.new("Outlook.Application") 
      end
      ns = @ol.GetNameSpace("MAPI")
      folder = ns.GetDefaultFolder(9) #olFolderCalendar
      raise "olFolderCalendar View not found" unless folder
      @events = OutlookEvents.new(folder)

    end
    def quit
      @ol.Quit if @created             # TODO
    end 
  end
  class OutlookEvents
    def initialize(folder)
      @calendar_items = folder.Items
      @calendar_items.Sort("[Start]")
      @calendar_items.IncludeRecurrences = true
    end
    def each
      @calendar_items.each do |event|
        begin
          yield OutlookEvent.new(event)
        rescue  StandardError => e 
         print 'X'
         $logger.error DateTime.now.to_s
         $logger.error e 
         $logger.error event
        end
      end
    end
  end
  class OutlookEvent
    # Note: The value of the minutes can be any arbitrary number of minutes between 5 minutes to 4 weeks. 
    GCAL_MAX_REMINDER = 40320  # 4 Weeks in minutes
    GCAL_MIN_REMINDER = 5
    APPOINTMENTTYPE = {'0' =>:appointment,
                                   '1' => :anniversary, 
                                   '2' => :all_day_event,
                                   '3' => :meeting,
                                   '4' => :reminder}
    attr_accessor :uid, 
                       :subject, 
                       :where, 
                       :start_time, :end_time,
                       :last_modified, #
                       :appointmenttype,
                       :content,
                       :repeats,
                       :alarm, :alarm_offset,
                       :required_names, :optional_names, :chair # todo ???
                       
    def initialize(outlook_event)
      fill_alarm(outlook_event)   #TODO
      
      # Notes Id
      @uid = outlook_event.GlobalAppointmentID
      
      # Subject
      if outlook_event.Subject
         @subject = outlook_event.Subject
      else
         @subject = ''
         $logger.warn 'no subject. uid: '+@uid
      end 
     
      # Room/Location
      @where = outlook_event.Location || ''
      
      # start date + time
      @start_time = outlook_event.Start.to_s
      
      # end date + time
      @end_time =  outlook_event.End.to_s
      
      # event type  # TODO -> remove
      @appointmenttype = APPOINTMENTTYPE['3'] # Meeting
      
      #p outlook_event.LastModificationTime.to_s
      @last_modified = DateTime.parse(outlook_event.LastModificationTime.to_s)
      @content =  outlook_event.Body

      # -- neue 
      @all_day_event = outlook_event.AllDayEvent

     # fill_repeats(outlook_event)
     
      @chair = outlook_event.Organizer
      @required_names = outlook_event.RequiredAttendees
      @optional_names = outlook_event.OptionalAttendees
    end
    def all_day?
      @all_day_event
    end
    def meeting?
      true # TODO
    end

    def supported?
      # anniversaries are now (v.0.0.7) supported 
      @appointmenttype #and (@appointmenttype != :anniversary)
    end
    def repeats?
      @repeats.size > 1
    end
    def fill_alarm(outlook_event)
      # Alarm
      @alarm = false
      @alarm_offset = 0
      if outlook_event.ReminderSet
      then
        @alarm = true
        @alarm_offset = outlook_event.ReminderMinutesBeforeStart
      end

    end
    def fill_repeats(outlook_event)
      @repeats = []
    end
    def formatted_names
      names = "Chair: #{@chair}\nRequired: "
      names += (@required_names || "")
      names += "\nOptional: "+ (@optional_names || "")
      names += "\n---\n"
      #p names
      names 
    end
    
    
    # obsolete methods -> TODO remove!!
    def _names_to_str_name_email(names)
      (names.inject([]) do |x,y| 
        email, name = '', ''
        email = "<#{y[:email]}>" if (y[:email] and y[:email] != '')
        name = "\"#{y[:name]}\" " if (y[:name] and y[:name] != '')

        x << name + email
        x 
      end).join(', ')
    end

    def _names_to_str(names)
      (names.inject([]) do |x,y| 
        if y[:email] and y[:email] != ''
           x << y[:email] if y[:email]
        elsif y[:name]
           x << y[:name] if y[:name]
        end
        x 
      end).join(', ')
    end
    def _all_names
      names = []
      names += @chair || [] 
      names += @required_names || []
      names += @optional_names || []
      
      names.uniq
    end
    def _fill_names(outlook_event)
      @chair = outlook_event.Organizer
      @required_names = outlook_event.RequiredAttendees
      @optional_names = outlook_event.OptionalAttendees
    end  
    def _fill_chair(outlook_event)
      chair = []
      chair << outlook_event.Organizer 
      chair
    end
    def _find_email(names, idx)
      email = names[idx]     
      if email 
        email = nil if email == '.'
        email = nil if email == ''
        # email =~ /^CN=(.*)\/O=(.*)\/C=(.*)/
        # email = $1 if $1
      end
      return email
    end
    
    def _fill_notes_names(outlook_event, notes_attr_name, notes_attr_email = nil)
      names = []
      notes_names = []
      notes_names1 = outlook_event.GetFirstItem(notes_attr_name)
      if notes_attr_email 
        notes_names2 = outlook_event.GetFirstItem(notes_attr_email) 
        notes_names2 = nil unless (notes_names2 and notes_names2.Values.size == notes_names1.Values.size)
      end
      if notes_names1
        notes_names1.Values.each_with_index do |name, idx|
            email = find_email(notes_names2.Values, idx) if notes_names2
            #name =~ /^CN=(.*)\/O=(.*)\/C=(.*)/
            #email ? short_name = name.split('/')[0] : short_name  # use name+domain if email missing
            short_name = name.split('/')[0]
            # check if name is an email adress
            if !email and short_name =~ /\A([^@\s]+)@((?:[-a-z0-9]+\.)+[a-z]{2,})\Z/i
               email = short_name
               short_name = ''
            end
            names << {:name => (short_name || ''), :email => (email || '')}
        end
      end
      names
    end
    
  end
end