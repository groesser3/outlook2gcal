# http://www.koders.com/ruby/fid0F7B2481BAB417B962DBED27C033DA0E00BEC6EB.aspx?s=proxy#L31

# http://msdn.microsoft.com/en-us/library/dd469461(v=office.12).aspx
# http://msdn.microsoft.com/en-us/library/ff870662.aspx

#encoding: utf-8
require "win32ole"

def each_event
  created = false
  ol = nil
  begin 
    ol = WIN32OLE.connect("Outlook.Application") 
  rescue 
    created = true
    ol = WIN32OLE.new("Outlook.Application") 
  end
  ns = ol.GetNameSpace("MAPI")
  myAppointments = ns.GetDefaultFolder(9).Items #olFolderCalendar
  myAppointments.Sort("[Start]")
  myAppointments.IncludeRecurrences = true
  myAppointments.each do |event|
    yield event
  end
  ol.Quit if created
end


# Delete All Future Data Of Google Calendar
now = Time.now
# Insert All Future Data Of Outlook
@nstr = now.strftime("%Y/%m/%d %H:%M:%S")
each_event do |oev|
#  if oev.End > @nstr
    p oev.Subject 
    
    #NKF is used for japanese charcter code conversion
   begin
    puts oev.Subject
    puts oev.Location
    puts oev.Start
    puts oev.End
    puts oev.AllDayEvent
    puts oev.IsRecurring
    puts oev.GlobalAppointmentID
    puts oev.LastModificationTime
    puts oev.RequiredAttendees
    puts oev.OptionalAttendees
   rescue
     puts "ERROR" + "*" *15     
   end 
    puts "-"*10
    
=begin
Notes-Attribute 
         :uid,    # GlobalAppointmentID 
         :subject,  #Subject
         :where, #Location
         :start_time, :end_time, # Start, End
         :last_modified, #LastModificationTime 
         :appointmenttype,  # '2' entspricht OL#AllDayEvent 
         :content,         # Body 
         :repeats,   #IsRecurring 
         :alarm, 
         :alarm_offset,
         :required_names, #RequiredAttendees 
         :optional_names, #OptionalAttendees 
         :chair # todo ???
         
         
    #EntryID  kann sich ändern
    #    Attachments
    #CreationTime 
    #Duration 
    #Organizer 
=end

#  end 
end