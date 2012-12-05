require "rubygems"

spec = Gem::Specification.new do |s|
  s.name = %q{outlook2gcal}
  s.version = "0.0.1"
  s.authors = ["Elias Kugler"]
  s.email = %q{groesser3@gmail.com}
  s.files =   Dir["lib/**/*"] + Dir["bin/**/*"] + Dir["*.rb"] + ["MIT-LICENSE","outlook2gcal.gemspec"]
  s.platform    = Gem::Platform::RUBY
  s.has_rdoc = true
  s.extra_rdoc_files = ["README.rdoc", "CHANGELOG.rdoc"]
  s.require_paths = ["lib"]
  s.summary = %q{Sync your Outlook calendar with your Google calendar}
  s.description = %q{This lib/tool syncs your Outlook calendar with your (private) Google calendar. The synchronisation is only one-way: Outlook events are pushed to your Google Calendar. All types of events (including recurring events like anniversaries) are supported.}
  s.files.reject! { |fn| fn.include? "CVS" }
  s.require_path = "lib"
  s.default_executable = %q{outlook2gcal}
  s.executables = ["outlook2gcal"]
  s.homepage = %q{http://rubyforge.org/projects/outlook2gcal/}
  s.rubyforge_project = %q{outlook2gcal}
  s.add_dependency("dm-core", ">= 1.2.0")
  s.add_dependency("dm-migrations", ">= 1.2.0")
  s.add_dependency("do_sqlite3", ">= 0.10.8")
  s.add_dependency("gdata4ruby", "=0.1.5")
  s.add_dependency("groesser3-gcal4ruby", "=0.5.51")  #!!! you need a modified version of this gem !!!
  s.add_dependency("log4r", ">=1.1.10")

end


