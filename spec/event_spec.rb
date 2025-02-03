# frozen_string_literal: true

require 'rspec'
require './event'
require 'vcalendar'
require 'json'
require 'yaml'

TEST_FILE = 'feed2.ics'
CONFIG_FILE = 'config.yml'

RSpec.describe 'Event' do
  before do
    ics = File.read TEST_FILE
    cal = Vcalendar.parse(ics, false)
    rcal = cal.to_hash
    @events = rcal[:VCALENDAR][:VEVENT]
    config = YAML.safe_load_file(CONFIG_FILE)
    @weekly_patterns = config['weekly']
  end

  after do
    # Do nothing
  end

  describe '#short_event_times' do
    it 'parses a Vcalendar event hash' do
      expect(Event.new(@events[0], @weekly_patterns).short_event_times).to eq('Friday 5 Jan 10:00 - 11:00')
    end
  end
end
