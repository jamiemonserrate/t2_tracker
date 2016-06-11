require 't2_generator'
require 'timecop'

describe T2Generator do
  it 'should generate file for all weeks' do
    Timecop.travel(Date.parse '05 April 2015')

    T2Generator.new('config/valid.yml').run

    workbook = RubyXL::Parser.parse('SomeName_T2_Tracker_Sheet.xlsx')

    expect(workbook.worksheets.size).to eq(5)

    ['02 March 2015', '09 March 2015', '16 March 2015', '23 March 2015'].each_with_index do |date, index|
      sheet = workbook.worksheets[index]
      expect(sheet.sheet_name).to eq(date)
      expect(sheet[3][1].value).to eq('Employee Name')
      expect(sheet[4][1].value).to eq(date)
      (8..17).each do |row_index|
        expect(sheet[row_index][2].value).to eq('Office Location')
      end
    end

    # Week containing Bank Holiday and Annual Leave
    sheet5 = workbook.worksheets[4]
    expect(sheet5[8][2].value).to eq('Office Location')
    expect(sheet5[9][2].value).to eq('Office Location')

    expect(sheet5[10][2].value).to eq('Office Location')
    expect(sheet5[11][2].value).to eq('Office Location')

    expect(sheet5[12][2].value).to eq('Annual Leave')
    expect(sheet5[13][2].value).to eq('Annual Leave')

    expect(sheet5[14][2].value).to eq('Office Location')
    expect(sheet5[15][2].value).to eq('Office Location')

    expect(sheet5[16][2].value).to eq('Bank Holiday')
    expect(sheet5[17][2].value).to eq('Bank Holiday')
  end

  it 'should raise error if start date is not a monday' do
    Timecop.travel(Date.parse '15 March 2015')

    expect(Proc.new{T2Generator.new('config/invalid_start_date.yml').run}).to raise_error("Please provide a start day that is a Monday")
  end

  after do
    FileUtils.rm_f('SomeName_T2_Tracker_Sheet.xlsx')
    Timecop.return
  end
end
