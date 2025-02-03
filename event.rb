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
  attr_reader :desc, :start, :end, :room, :url

  def initialize(event_hash, weekly_patterns)
    @start = vtparse(event_hash[:DTSTART])
    @end = vtparse(event_hash[:DTEND])
    @desc = event_hash[:SUMMARY][:value]
    @weekly = false
    weekly_patterns.each do |pstr|
      pattern = Regexp.new(pstr, 'i')
      @weekly = true if pattern.match? desc
    end
    descstr = event_hash[:DESCRIPTION][:value].split(/\n/)
    @room = descstr[2]
    @url = event_hash[:URL][:value]
  end

  def short_event_times
    "#{@start.strftime('%A %-d %b %H:%M')} - #{@end.strftime('%H:%M')}"
  end
end
