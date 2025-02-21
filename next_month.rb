# frozen_string_literal: true

# Usage: "ruby next_month.rb /tmp/events-%.xlsx" outputs /tmp/events-2025-03.xlsx
# This is a utility script which might be useful in naming the output file from vh_events.rb.
require 'date'
ymstr = (Date.today >> 1).strftime('%Y-%m')
puts ARGV[0].gsub('%', ymstr)
