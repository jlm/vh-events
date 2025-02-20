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
class TemplateParser
  attr_reader :dayrow, :daycol, :roomcount, :roomcolumns, :worksheet, :workbook, :debug

  def initialize(template_filename, days, debug: false)
    @days = days
    @debug = debug
    begin
      @workbook = RubyXL::Parser.parse(template_filename)
    rescue Zip::Error => e
      abort e.message
    end
    @worksheet = @workbook[0]
    parse_template(@worksheet)
  end

  def find_monthly_section
    ((@dayrow + 1)..(@dayrow + 8)).each do |rownumber|
      next if @days.values.include?(@worksheet[rownumber][@daycol].value)

      return rownumber
    end
    nil
  end

  def parse_config
    configws = @workbook['Config'] || raise('No Config sheet found in template')
    configrow = nil
    configws.each_with_object([]) do |row, result|
      if /name in feed/i.match?(row[0].value)
        configrow = row.r - 1
        next
      end
      next unless configrow

      result << make_rule(row)
    end
  end

  def make_rule(row)
    { 'pattern' => row[0]&.value,
      'pub_name' => row[1]&.value,
      'weekly' => /ye?s?/i.match(row[2]&.value) || /true/i.match(row[2]&.value),
      'time_rule' => row[3]&.value,
      'termtime' => /ye?s?/i.match(row[4]&.value) || /true/i.match(row[4]&.value) }
  end
  private :make_rule

  def new_cell(row, column, content, wrap: true, weight: 'thin')
    worksheet.add_cell(row, column, content)
    cell = worksheet[row][column]
    cell.change_text_wrap(wrap)
    cell.change_text_indent(1)
    cell.change_border(:top, weight)
    cell.change_border(:bottom, weight)
    cell.change_border(:left, weight)
    cell.change_border(:right, weight)
  end

  def add_monthly_events_to_worksheet(other_events_by_date, other_events_startline)
    (0..(other_events_by_date.length - 1)).each do |ev_index|
      rownumber = other_events_startline + 1 + ev_index
      oe_entry = other_events_by_date[ev_index]
      date = oe_entry[:date]
      # Mon 3rd Feb
      datestring = date.strftime('%a ') + date.day.ordinalize + date.strftime(' %b')
      new_cell(rownumber, daycol, datestring, wrap: false)
      event_descs = {}
      oe_entry[:events].each_value do |pevs|
        pevs.each do |event|
          event.room.split(', ').each do |room|
            puts "#{rownumber} #{datestring} #{room} #{event.desc_and_times}" if debug
            event_descs[room] ||= []
            event_descs[room] << event.desc_and_times
          end
        end
      end
      event_descs.each do |room, event_descriptions|
        new_cell(rownumber, roomcolumns[room], event_descriptions.join("\n"))
      end
    end
  end

  def add_weekly_events_to_worksheet(wkt_by_day)
    wkt_by_day.each do |entry|
      ((dayrow + 1)..(dayrow + 8)).each do |rownumber|
        next unless worksheet[rownumber][daycol].value == entry[:day]

        entry[:events].each do |room, pevs|
          new_cell(rownumber, roomcolumns[room], pevs.map(&:desc_and_times).join("\n"))
        end
      end
    end
  end

  def output_events_as_xlsx(wkt_by_day, other_events_by_date, excel_filename)
    add_weekly_events_to_worksheet(wkt_by_day)
    other_events_startline = find_monthly_section || raise("Couldn't find monthly section in template file")
    add_monthly_events_to_worksheet(other_events_by_date, other_events_startline)
    workbook.write(excel_filename)
  end

  private

  def parse_template(worksheet)
    @dayrow = nil # Excel row number of the row starting with "Day"
    @daycol = nil # Excel column number of the column containing "Day"
    @roomcount = nil
    worksheet.each do |row|
      (0..(row.size - 1)).each do |column|
        cell = row[column]
        if cell&.value == 'Day'
          @dayrow = row.r - 1
          @daycol = column
          @roomcount = row.size - 1 - @daycol
          puts "Day cell: #{cell.r}" if @debug
          break
        end
        break if @daycol
      end
    end
    raise '"Day" cell not found in template spreadsheet' if @daycol.nil?

    @roomcolumns = {}
    ((@daycol + 1)..(@daycol + @roomcount)).each do |column|
      @roomcolumns[worksheet[@dayrow][column]&.value] = column
    end
  end
end
