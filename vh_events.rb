# frozen_string_literal: true

####
# Copyright 2025 John Messenger
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
####
require 'rubygems'
require 'bundler/setup'
Bundler.require

require 'vcalendar'
require 'json'
require 'active_support/core_ext/integer/inflections'
require 'active_support/deprecation'
require 'active_support/deprecator'
require './event'
require './template_parser'
require 'rubyXL/convenience_methods/workbook'
require 'rubyXL/convenience_methods/cell'

DAYS = { 0 => 'Sunday', 1 => 'Monday', 2 => 'Tuesday', 3 => 'Wednesday', 4 => 'Thursday', 5 => 'Friday',
         6 => 'Saturday' }.freeze

def print_events(events, title)
  puts title
  events.each do |pev|
    puts pev
  end
end

# This program parses a Vcalendar feed of Village Hall events produced by Hallmaster, and produces
# an Excel spreadsheet for inclusion in the Magazine.  A configuration file (config.yml) allows events to be
# categorized as weekly (with others being designated non-weekly), and to be recognised using a regular expression
# rather than an exact text match.  It also allows events to be renamed in the output table. (Although Hallmaster
# has a "weekly" event field, many events in the database which are weekly are not marked as such.)
# The spreadsheet is based on a template spreadsheet provided as input,
# which can be hand-crafted in Excel to include appropriate styling, borders, day names and room names.
# These are read from the spreadsheet template rather than only from the event feed, to provide configurability
# of the output. For example, if a row were added to the template for weekly events on a Saturday, then the program
# would include such events in the table, instead of ignoring them as it does at present.
#
# Parse command line options
begin
  opts = Slop.parse do |o|
    o.string '-c', '--config', 'configuration YAML file name (default: config.yml)', default: 'config.yml'
    o.bool '-C', '--config-from-template', 'take rewrite rules from the template file'
    o.string '-t', '--template', 'name of the Excel template file to read'
    o.string '-r', '--read-file', 'Read iCal data from file (default: read from URL in config file)'
    o.integer '-m', '--month', 'the (numeric) month to search (default: next month)'
    o.integer '-y', '--year', 'the year to search'
    o.bool '-d', '--debug', 'debug mode'
    o.bool '-v', '--verbose', 'be verbose: list extra detail'
    o.bool '-a', '--print-all', 'print all events'
    o.string '-j', '--json', 'output results in JSON to the named file'
    o.string '-x', '--excel', 'output results in XLSX to the named file'
    o.bool '-e', '--email', 'email the output file to the address in config file'
    o.bool '-l', '--list', 'list events to standard output'
    o.on '--help' do
      warn o
      exit
    end
  end

  # Read additional configuration from the config file.
  config = YAML.safe_load_file(opts[:config])
  # Open the JSON output file, if used.  XXX: JSON output only includes weekly events
  jfile = opts[:json] ? File.open(opts[:json], 'w') : nil
  # Process Excel options and open Excel template file
  puts "Writing Excel output to #{opts[:excel]}" if opts[:excel] && opts.verbose?
  template_name = opts[:template] || (config['excel'] ? config['excel']['template'] : nil)
  template = nil
  template = TemplateParser.new(template_name, DAYS, debug: opts.debug?) if template_name
  abort('Excel output selected but no template provided') if template.nil? && opts[:excel]
  # Read the event parsing rules
  if template_name && (opts.config_from_template? || (config['excel'] && config['excel']['config_from_template']))
    puts 'Reading parse_rules from the Excel template file' if opts.verbose?
    parse_config = template&.parse_config
  else
    puts "Reading parse_rules from #{opts['config']}" if opts.verbose?
    parse_config = config['parse_rules']
  end
  warn 'No parsing configuration rules found' if parse_config.nil?

  # Process event selection command-line options
  # If month is specified and is zero, select all events rather than just that month's events.
  # NB: could just check for month == 0.
  select_all = opts[:month] && (opts[:month]).zero?
  # By default, use next month and its year
  month = opts[:month] || (Date.today >> 1).month
  year = opts[:year] || (Date.today >> 1).year
  window_start = Date.new(year, month, 1).to_time unless select_all
  window_end = (Date.new(year, month, 1) >> 1).to_time unless select_all
  puts "Selected month: #{year} #{month}" if opts.debug?

  # Obtain and parse the event feed or file
  ics = (opts[:read_file] ? File.open(opts[:read_file]) : URI.open(config['url'])).read.gsub(/\r\n/, "\n")
  vevents = Vcalendar.parse(ics, false).to_hash[:VCALENDAR][:VEVENT]
  parsed_events = vevents.map do |event|
    Event.new(event, parse_config)
  end
  # Select the desired events
  parsed_events.sort_by!(&:start)
  print_events(parsed_events, 'All events in feed') if opts.print_all?
  unless select_all
    parsed_events.reject! do |pev|
      pev.start < window_start || pev.end > window_end
    end
  end
  print_events(parsed_events, 'Selected events') if opts.list?

  weekly_events = parsed_events.select(&:weekly)
  other_events = parsed_events.reject(&:weekly)

  if opts.verbose?
    puts "* Number of events: #{parsed_events.size}. Weekly: #{weekly_events.size}. Other Events: #{other_events.size}"
  end

  # Uniquify the weekly events, sort and group by day-of-week
  wkt = weekly_events.uniq(&:day_desc_and_times).sort_by(&:day).group_by(&:day)
  # Make a table of day-of-week with columns for each room
  wkt_by_day = wkt.map do |day, pevs|
    { day: DAYS[day], events: pevs.sort_by(&:start).group_by(&:room) }
  end

  # Group the remaining (non-weekly) events by date, and make a table of them with columns for each room
  ott = other_events.group_by(&:date).map do |date, pevs|
    { date: date, events: pevs.sort_by(&:start).group_by(&:room) }
  end

  # Process the outputs, saving to files or sending by email
  if opts[:excel]
    filename = opts[:excel]
    template.output_events_as_xlsx(wkt_by_day, ott, filename)
    if opts[:email]
      mail_from = config['mail_from'] || ENV['MAIL_FROM']
      mail_to = (config['mail_to'] || ENV['MAIL_TO']).split(',').map(&:strip)
      postmark_api_key = config['postmark_api_key'] || ENV['POSTMARK_API_KEY']

      message = Mail.new do
        from            mail_from
        to              mail_to
        subject         "Events table for #{year}-#{month}"
        body            "Please find enclosed the auto-generated events table.\n"

        delivery_method Mail::Postmark, api_token: postmark_api_key
      end

      _dir, rest = File.split(filename)
      message.attachments[rest] = File.read(filename)
      message.deliver
      if message.delivered
        puts "Message #{message.postmark_response['MessageID']} delivered to #{mail_to.size} recipients" if opts.verbose?
      else
        warn 'Message not delivered'
      end
    end
  end
  jfile&.puts JSON.pretty_generate(wkt_by_day)
  # print_events(weekly_events, 'Weekly Events')
end
