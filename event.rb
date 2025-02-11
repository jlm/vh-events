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

  def initialize(event_hash, weekly_patterns)
    @start = vtparse(event_hash[:DTSTART])
    @day = @start.wday
    @end = vtparse(event_hash[:DTEND])
    @desc = event_hash[:SUMMARY][:value]
    @weekly = false
    weekly_patterns.each do |pstr|
      (patternstr, substitute) = pstr.split('|')
      pattern = Regexp.new(patternstr, 'i')
      if pattern.match? desc
        @weekly = true
        @desc = substitute if substitute
      end
    end
    descstr = event_hash[:DESCRIPTION][:value].split(/\n/)
    @room = descstr[2]
    @url = event_hash[:URL][:value]
  end

  def short_event_times
    "#{@start.strftime('%A %-d %b %H:%M')} - #{@end.strftime('%H:%M')}"
  end

  def desc_and_times
    "#{@desc} #{@start.strftime('%H:%M')} - #{@end.strftime('%H:%M')}"
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
