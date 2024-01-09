require 'rubyXL'

class ExcelController < ApplicationController
  def read_excel
    workbook = RubyXL::Parser.parse('C:/Nueva carpeta/file.xlsx')
    worksheet = workbook[0]

    @data = []
    worksheet.each do |row|
      row_data = row.cells.map { |cell| cell&.value } 
      @data << row_data unless row_data.compact.empty?
    end
  end

  def write_excel_form
    
  end

  def write_excel
    workbook = RubyXL::Parser.parse('C:/Nueva carpeta/file.xlsx')
    worksheet = workbook[0]

    row = params[:row].to_i
    column = params[:column].to_i
    value = params[:value]

    if row >= 0 && column >= 0
      worksheet.add_cell(row, column, value)
      workbook.write('C:/Nueva carpeta/modified_file.xlsx')
    end

    redirect_to action: :read_excel 
  end
end
