# frozen_string_literal: true

# Usage: "ruby next_month.rb /tmp/events-%.xlsx" outputs /tmp/events-2025-03.xlsx
require 'date'
ymstr = (Date.today >> 1).strftime('%Y-%m')
puts ARGV[0].gsub('%', ymstr)
