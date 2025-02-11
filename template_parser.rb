# frozen_string_literal: true

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
