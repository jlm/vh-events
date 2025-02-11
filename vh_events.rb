# frozen_string_literal: true

####
# Copyright 2016-2022 John Messenger
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
require './event'
require './template_parser'

DAYS = { 0 => 'Sunday', 1 => 'Monday', 2 => 'Tuesday', 3 => 'Wednesday', 4 => 'Thursday', 5 => 'Friday',
         6 => 'Saturday' }.freeze

def print_events(events, title)
  puts title
  events.each do |pev|
    puts pev
  end
end

def add_monthly_events_to_worksheet(template, other_events_by_date, other_events_startline)
  (0..(other_events_by_date.length - 1)).each do |ev_index|
    rownumber = other_events_startline + 1 + ev_index
    oe_entry = other_events_by_date[ev_index]
    date = oe_entry[:date]
    datestring = date.strftime('%a ') + date.day.ordinalize + date.strftime(' %b')
    template.worksheet.add_cell(rownumber, template.daycol, datestring)
    event_descs = {}
    oe_entry[:events].each_value do |pevs|
      pevs.each do |event|
        event.room.split(', ').each do |room|
          puts "#{rownumber} #{datestring} #{room} #{event.desc_and_times}" if template.debug
          event_descs[room] ||= []
          event_descs[room] << event.desc_and_times
        end
      end
    end
    event_descs.each do |room, event_descriptions|
      template.worksheet.add_cell(rownumber, template.roomcolumns[room], event_descriptions.join("\n"))
    end
  end
end

def add_weekly_events_to_worksheet(template, wkt_by_day)
  wkt_by_day.each do |entry|
    ((template.dayrow + 1)..(template.dayrow + 8)).each do |rownumber|
      next unless template.worksheet[rownumber][template.daycol].value == entry[:day]

      entry[:events].each do |room, pevs|
        template.worksheet.add_cell(rownumber, template.roomcolumns[room], pevs.map(&:desc_and_times).join("\n"))
      end
    end
  end
end

def output_events_as_xlsx(wkt_by_day, other_events_by_date, excel_filename, excel_conf, debug)
  template = TemplateParser.new(excel_conf['template'], DAYS, debug: debug)

  add_weekly_events_to_worksheet(template, wkt_by_day)
  other_events_startline = template.find_monthly_section
  raise "Couldn't find start of monthly section" unless other_events_startline

  add_monthly_events_to_worksheet(template, other_events_by_date, other_events_startline)

  template.workbook.write(excel_filename)
end

begin
  opts = Slop.parse do |o|
    o.string '-c', '--config', 'configuration YAML file name', default: 'config.yml'
    o.string '-r', '--read-file', 'Read iCal data from file'
    o.integer '-y', '--year', 'the year to search'
    o.integer '-m', '--month', 'the month to search'
    o.bool '-d', '--debug', 'debug mode'
    o.bool '-a', '--print-all', 'print all selected events'
    o.bool '-v', '--verbose', 'be verbose: list extra detail'
    o.string '-j', '--json', 'output results in JSON to the named file'
    o.string '-x', '--excel', 'output results in XLSX to the named file'
    o.bool '-l', '--list', 'list events'
    o.on '--help' do
      warn o
      exit
    end
  end

  config = YAML.safe_load_file(opts[:config])
  jfile = opts[:json] ? File.open(opts[:json], 'w') : $stdout
  puts "Writing Excel output to #{opts[:excel]}" if opts[:excel] && opts.verbose?

  # If month is specified and is zero, select all events rather than just that month's events.
  # NB: could just check for month == 0.
  select_all = opts[:month] && (opts[:month]).zero?
  month = opts[:month] || (Date.new >> 2).strftime('%m').to_i
  year = opts[:year] || Time.now.strftime('%Y').to_i
  window_start = Date.new(year, month, 1).to_time
  window_end = (Date.new(year, month, 1) >> 1).to_time
  puts "Selected month: #{year} #{month}" if opts.debug?

  ics = nil
  ics = File.read opts[:read_file] if opts[:read_file]
  cal = Vcalendar.parse(ics, false)
  rcal = cal.to_hash
  events = rcal[:VCALENDAR][:VEVENT]

  parsed_events = events.map do |event|
    Event.new(event, config['weekly'])
  end
  parsed_events.sort_by!(&:start)
  print_events(parsed_events, 'Events in window') if opts.print_all?
  unless select_all
    parsed_events.reject! do |pev|
      pev.start < window_start || pev.end > window_end
    end
  end

  weekly_events = parsed_events.select(&:weekly)
  other_events = parsed_events.reject(&:weekly)

  if opts.verbose?
    puts "* Number of events: #{parsed_events.size}. Weekly: #{weekly_events.size}. Other Events: #{other_events.size}"
  end

  wkt = weekly_events.sort_by(&:day).sort_by(&:start).uniq(&:desc_and_times).sort_by(&:day).group_by(&:day)
  wkt_by_day = wkt.map do |day, pevs|
    { day: DAYS[day], events: pevs.sort_by(&:start).group_by(&:room) }
  end

  ott = other_events.group_by(&:date).map do |date, pevs|
    { date: date, events: pevs.sort_by(&:start).group_by(&:room) }
  end

  jfile.puts JSON.pretty_generate(wkt_by_day) if opts.json?
  output_events_as_xlsx(wkt_by_day, ott, opts[:excel], config['excel'], opts.debug?) if opts[:excel]

  # print_events(weekly_events, 'Weekly Events')
end
