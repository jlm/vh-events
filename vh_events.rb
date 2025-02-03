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

class TimeWithTimezone < Time
  require 'timezone'

  def self.find_timezone(zone)
    Timezone[zone]
  end
end

def vtparse(evttime)
  tt = evttime[:value]
  ts = "#{tt[:year]}-#{tt[:month]}-#{tt[:day]} #{tt[:hour]}:#{tt[:min]}:#{tt[:sec]} #{tt[:zone]}"
  TimeWithTimezone.parse(ts)
end

def short_event_times(start, finish)
  "#{start.strftime('%A %-d %b %H:%M')} - #{finish.strftime('%H:%M')}"
end

def desc_and_times(event)
  "#{event[:desc]} #{event[:start].strftime('%H:%M')} - #{event[:end].strftime('%H:%M')}"
end

def print_events(events, title)
  puts title
  events.each do |pev|
    puts "#{pev[:desc]} #{pev[:room]} #{pev[:short_event_times]}"
  end
end

begin
  opts = Slop.parse do |o|
    o.string '-c', '--config', 'configuration YAML file name', default: 'config.yml'
    o.string '-r', '--read-file', 'Read iCal data from file'
    o.integer '-y', '--year', 'the year to search'
    o.integer '-m', '--month', 'the month to search'
    o.bool '-d', '--debug', 'debug mode'
    o.bool '--print-all', 'print all selected events'
    o.bool '-v', '--verbose', 'be verbose: list extra detail (unimplemented)'
    o.bool '-j', '--json', 'output results in JSON'
    o.bool '--slackpost', 'post alerts to Slack for new items'
    o.bool '-l', '--list', 'list events'
    o.bool '--list-sizes', 'List available slug sizes'
    o.bool '--list-images', 'List available OS images'
    o.string '--region', 'region to create the vm in'
    o.string '--size', 'size_slug of the vm to be created'
    o.string '--ipv6', 'enable ipv6?'
    o.string '--image', 'image name of the vm to be created'
    o.string '--tags', 'list of tags to tag the vm with'
    o.bool '--destroy', 'destroy the vm'
    o.string '--dhdns-add', 'add the vm to the given domain in Dreamhost DNS'
    o.string '--dhdns-remove', 'remove the Dreamhost DNS name of the vm from the given domain'
    o.bool '--force', 'try to force the operation despite errors'
    o.on '--help' do
      warn o
      exit
    end
  end

  config = YAML.safe_load_file(opts[:config])

  month = opts[:month] || Time.now.strftime('%m').to_i
  year = opts[:year] || Time.now.strftime('%Y').to_i
  window_start = Date.new(year, month, 1).to_time
  window_end = (Date.new(year, month, 1) >> 1).to_time
  puts "Selected month: #{year} #{month}"
  ics = nil
  ics = File.read opts[:read_file] if opts[:read_file]
  cal = Vcalendar.parse(ics, false)
  rcal = cal.to_hash
  events = rcal[:VCALENDAR][:VEVENT]

  parsed_events = events.map do |event|
    starttime = vtparse(event[:DTSTART])
    endtime = vtparse(event[:DTEND])
    desc = event[:SUMMARY][:value]
    weekly = false
    config['weekly'].each do |pstr|
      pattern = Regexp.new(pstr, 'i')
      weekly = true if pattern.match? desc
    end
    descstr = event[:DESCRIPTION][:value].split(/\n/)
    room = descstr[2]
    url = event[:URL][:value]
    {
      desc: desc,
      start: starttime,
      end: endtime,
      short_event_times: short_event_times(starttime, endtime),
      room: room,
      weekly: weekly,
      url: url
    }
  end

  parsed_events.sort_by! do |pev|
    pev[:start]
  end

  unless opts.print_all?
    parsed_events.reject! do |pev|
      pev[:start] < window_start || pev[:end] > window_end
    end
  end

  print_events(parsed_events, 'Events in window') if opts.print_all?

  weekly_events = parsed_events.select do |pev|
    pev[:weekly]
  end

  other_events = parsed_events.reject do |pev|
    pev[:weekly]
  end

  puts "**** Number of events: #{parsed_events.size}. Weekly: #{weekly_events.size}. Other Events: #{other_events.size}"

  DAYS = { 0 => 'Sunday', 1 => 'Monday', 2 => 'Tuesday', 3 => 'Wednesday', 4 => 'Thursday', 5 => 'Friday',
           6 => 'Saturday' }.freeze
  weekly_events_by_days = DAYS.map do |daynum, _dayname|
    weekly_events.select do |pev|
      pev[:start].wday == daynum
    end
  end
  _james = 3
  DAYS.each do |daynum, dayname|
    puts "#{dayname}:"
    weekly_events_by_days[daynum].each do |pev|
      puts "    #{desc_and_times(pev)}"
    end
  end

  # print_events(weekly_events, 'Weekly Events')
end
