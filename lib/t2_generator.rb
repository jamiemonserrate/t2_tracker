require 'yaml'
require 'rubyXL'

require 'holidays'
require 'holidays/core_extensions/date'
class Date
  include Holidays::CoreExtensions::Date
end

class T2Generator
  MONDAY_START_ROW = 8
  BANK_HOLIDAY = 'Bank Holiday'
  ANNUAL_LEAVE = 'Annual Leave'

  def initialize(config_path)
    @config = YAML.load_file(config_path)
    @leaves = @config['leaves'].collect {|leave| Date.parse(leave)}
  end

  def run
    book = RubyXL::Parser.parse(template_file_name)
    week_commencing_date = Date.parse(@config['start_date'])

    book.worksheets = []

    while Date.today > week_commencing_date
      book.worksheets << generate_sheet_for(week_commencing_date)
      week_commencing_date += 7
    end

    book.write @config['output_file_name']
  end

  private

  def generate_sheet_for(date)
    raise "Please provide a start day that is a Monday" unless date.monday?

    sheet = RubyXL::Parser.parse(template_file_name)[0]

    sheet.sheet_name = format date
    sheet[3][1].change_contents(@config['name'])
    sheet[4][1].change_contents(format(date))

    5.times do |index|
      row_date = date + index
      if row_date.holiday?(:gb)
        value = BANK_HOLIDAY
      elsif @leaves.include?(row_date)
        value = ANNUAL_LEAVE
      else
        value = @config['office']
      end

      # since the sheet has AM and PM as seperate rows
      row_offset = MONDAY_START_ROW + 2 * index
      sheet[row_offset][2].change_contents(value)
      sheet[row_offset+1][2].change_contents(value)
    end

    return sheet
  end

  def format(date)
    date.strftime("%d %B %Y")
  end

  def template_file_name
    'template/new.xlsx'
  end
end
