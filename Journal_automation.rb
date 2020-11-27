#!/usr/bin/env ruby
require 'win32ole'
require 'csv'
require 'fileutils'

class Poop
  def get_masterdir(prompt='', title='')
    excel = WIN32OLE.new('Excel.Application')
    response = excel.InputBox(prompt, title)
    excel.Quit
    excel = nil
    return response
  end

  def get_issue(prompt='', title='')
    excel = WIN32OLE.new('Excel.Application')
    lst = excel.InputBox(prompt, title)
    excel.Quit
    excel = nil
    return lst
  end

  def get_pa(prompt='', title='')
    excel = WIN32OLE.new('Excel.Application')
    pa = excel.InputBox(prompt, title)
    excel.Quit
    excel = nil
    return pa
  end

end

  dir1 = Poop.new
  dir2 = Poop.new
  dir3 = Poop.new
  response = dir1.get_masterdir('enter directory of Master and JournalsLoadingTemplate. Must be in same dir. ex: C:\Journals', 'My Title')
  lst = dir2.get_issue('enter directory of .lst files. Ex: C:\issues', 'lst dir')
  pa = dir3.get_pa('enter name of PA in Master spreadsheet to process', 'PA Name')

  #output to log file
$stdout = File.open lst + "\\" + "output.log", "a"

xl = WIN32OLE.new('excel.application')
xl.Visible = false #hide excel instance
xl.ScreenUpdating = false #turn off screen updating
book = xl.workbooks.open(response+'\Master.xlsx')
xl.displayalerts = false #turn off all those pesky alert boxes,


book.worksheets.each do |sheet| 
  last_row = sheet.cells.find(what: '*', searchorder: 1, searchdirection: 2).row 
  last_col = sheet.cells.find(what: '*', searchorder: 2, searchdirection: 2).column
  export = File.new(response + '\\' + sheet.name + '.csv', 'w+')
  csv_row = []

 
  (1..last_row).each do |xlrow|
    (1..last_col).each do |xlcol|
      csv_row << sheet.cells(xlrow, xlcol).value
    end
    export << CSV.generate_line(csv_row)
    csv_row = []
  end
  export.close
end

# clean up
book.close(savechanges: 'false')
xl.quit




listFolder = lst + "\\"
outpath = response + "\\" + "OutputJournals" + "\\"

Dir.chdir(response)
unless Dir.exists?("OutputJournals")
Dir.mkdir("OutputJournals") 
end


excel = WIN32OLE.new('Excel.Application')
excel.DisplayAlerts = false
excel.Visible = false #hide excel instance
excel.ScreenUpdating = false #turn off screen updating


masterRow = 1 #track row number for master sheet
CSV.foreach(response + "\\" + 'Sheet1.csv') do |row|

if row.include?(pa) == true
then
puts "Production analyst found..."
puts "Creating..." + outpath + row[1] + ".xlsx...."
FileUtils.cp response + "\\" + "NewJournalsLoadingTemplate.xlsx", outpath + row[1] + ".xlsx"
listFile = row[0] + ".lst"
  list = listFolder+listFile
  if File.exist?(list) == true 
  then
    issue = File.readlines(list) 
    
    workbook = excel.Workbooks.Open(outpath + row[1] + ".xlsx")
    masterbook = excel.Workbooks.Open(response + "\\"+"Master.xlsx")
    worksheet1 = workbook.Worksheets(1)
    worksheet2 = workbook.Worksheets(2)
    mastersheet = masterbook.Worksheets("Sheet1")
    
    puts "Updating " + outpath + row[1] + ".xlsx......."
    
    worksheet1.Cells(2,2).Value = "#{row[1]} #{worksheet1.Cells(2,2).Value}"
    
    worksheet2.Cells(4,2).Value = row[1] #shortcode
    worksheet2.Cells(6,2).Value = worksheet2.Cells(6,2).Value + row[1] #Input file location
    worksheet2.Cells(3,2).Value = row[3] #journaltitle
    worksheet2.Cells(3,6).Value = row[0] #serial code
    worksheet2.Cells(4,6).Value = row[4] #issn
    worksheet2.Cells(6,6).Value = issue.length #total no of files/issues
    puts "Update done."
    
    puts "Updating " + response + "\\"+"Master.xlsx....."
    mastersheet.Cells(masterRow,7).Value = issue.length
    puts "Update Done."
    
    
    nrow = 10
    issue.each do |data|
    worksheet2.Cells(nrow,2).Value = data.chomp
    nrow+=1
    end
    
    workbook.Save 
    workbook.Close 
    
    masterbook.Save
    masterbook.Close
    
  else #else do nothing if file does not exist
    
  end #end of file.exist
else#pa is not found

end

masterRow+=1

end
excel.Quit 

 File.delete("C:\\rubyProject\\Sheet1.csv")

$stdout.close
$stdout = STDOUT

puts "Done"
