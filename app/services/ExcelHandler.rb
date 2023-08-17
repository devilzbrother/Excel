# app/services/excel_handler.rb

class ExcelHandler
  require 'roo'
  require 'axlsx'

  def self.read_excel(file_path)
    workbook = Roo::Spreadsheet.open(file_path)
    worksheet = workbook.sheet(0) # Assuming you want to read the first sheet

    data = []
    (2..worksheet.last_row).each do |row|
      row_data = worksheet.row(row)
      data << row_data
    end

    data
  end

  def self.write_excel(data, file_path)
    Axlsx::Package.new do |p|
      p.workbook.add_worksheet(name: 'Sheet1') do |sheet|
        data.each do |row|
          sheet.add_row(row)
        end
      end
      p.serialize(file_path)
    end
  end
end