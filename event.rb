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
require 'bundler'
Bundler.setup(:default)
require 'active_support/core_ext/time'

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

class Event
  attr_reader :desc, :start, :day, :end, :room, :url, :weekly

  include Comparable
  def <=>(other)
    start <=> other.start
  end

  # Given a time, parse new_time_string and extract just the hour and minute.  Return a time object based on the
  # time provided, but with the hour and minute replaced by those from the new_time_string.
  def change_time(a_time, new_time_string)
    new_time = Time.parse(new_time_string)
    a_time.change(hour: new_time.hour, min: new_time.min)
  end

  def initialize(event_hash, parsed_config)
    @start = vtparse(event_hash[:DTSTART])
    @day = @start.wday
    @end = vtparse(event_hash[:DTEND])
    @desc = event_hash[:SUMMARY][:value]
    descstr = event_hash[:DESCRIPTION][:value].split(/\n/)
    @room = descstr[2]
    @url = event_hash[:URL][:value]
    @weekly = false
    parsed_config&.each do |rule|
      pattern = Regexp.new(rule['pattern'], 'i')
      next unless pattern.match? desc

      @weekly = rule['weekly']
      @desc = rule['pub'] if rule['pub']
      @time_rule = rule['time_rule']
      case @time_rule
      when nil
        # nothing to do
      when /^=(\d+:\d+)(-(\d+:\d+))?/
        # Change @start and maybe @end times to be the same day but with the times set as per @time_rule
        @start = change_time(@start, ::Regexp.last_match(1))
        @end = change_time(@end, ::Regexp.last_match(3)) if ::Regexp.last_match(3)
        _james = 12
      when /\+(\d+)(-(\d+))?/
        @start += 60 * ::Regexp.last_match(1).to_i
        @end += 60 * ::Regexp.last_match(2).to_i if ::Regexp.last_match(2) # NOTE: this would be end = end + 60 * (-30)
      when /-(\d+)/
        @end -= 60 * ::Regexp.last_match(1).to_i
      else
        puts "Warning: can't parse time_rule: #{@time_rule}"
      end
    end
  end

  def short_event_times
    "#{@start.strftime('%A %-d %b %H:%M')} - #{@end.strftime('%H:%M')}"
  end

  def desc_and_times
    "#{@desc} #{@start.strftime('%H:%M')} - #{@end.strftime('%H:%M')}"
  end

  def day_desc_and_times
    "#{@start.strftime('%a')} #{@desc} #{@start.strftime('%H:%M')} - #{@end.strftime('%H:%M')}"
  end

  def date
    @start.to_date
  end

  def to_slong
    "#{@desc} #{@room} #{short_event_times}"
  end

  def to_s
    "#{@start.to_date} #{desc_and_times}"
  end
end
