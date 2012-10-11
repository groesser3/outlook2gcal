module Outlook2GCal

  class Counter
    attr_accessor :updates, :inserts, :deletes, :selects, :ignored, :t_start, :t_end
    
    def initialize
      @updates = 0
      @inserts = 0
      @deletes = 0
      @selects = 0
      @ignored = 0      
      start
    end
    def runtime
      @t_end - @t_start 
    end 
    def start
      @t_start = Time.now
    end
    def end
      @t_end = Time.now
    end
    def show
      puts "\nStatistics:"
      puts "-"*20
      puts sprintf("Inserts    : %05d", @inserts)
      puts sprintf("Updates    : %05d", @updates)
      puts sprintf("Deletes    : %05d", @deletes)
      puts sprintf("Ingored    : %05d", @ignored)
      puts "-"*20
      puts sprintf("Total      : %05d", @selects)
      puts          "Runtime    : #{runtime} sec"
    end
  end
  
end