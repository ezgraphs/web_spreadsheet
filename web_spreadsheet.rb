=begin
  web_spreadsheet.rb - Display an Excel Spreadsheet using OLE Automation through a sinatra web application.
    Copyright (C) 2010  EzGraphs LLC.

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    Questions or Concerns: ezgraphs  [at} g m a i l d0t C0m
=end
#
# Should be obvious at a glance: this only runs on Windows machines with Excel installed.
#

['sinatra', 'win32ole'].each{|r| require r}

# TODO: Make this configurable 
DIRECTORY = Dir.pwd

def err(msg)
  "<p style=\"color: red; font-weight: bold\">ERROR: #{msg}</p>"
end

#
#  Have had problems with excel processes hanging around.  Tried a variety
# of other solutions unsucessfully and resorted to killing all excel processes
#
def kill_excel_automation_processes
  wmi = WIN32OLE.connect("winmgmts://")
  processes = wmi.ExecQuery("select * from win32_process where commandline like '%excel.exe\"% /automation %'")
  processes.each{|process| Process.kill('KILL', process.ProcessID)}  
end 

def _header(msg,params)
   str="<h1>#{msg}</h1>"
   str+="<a href=\"/\">Spreadsheet List</a><br/>"
   unless params.nil?
     str+='<h3>Parameters</h3><table border="1" cellpadding="2" cellspacing="0">'
     params.each_pair{|k,v|str+="<tr><th align='left'>#{k}</th><td>#{v}</td></tr> \n"}
     str+='</table><br/>'
   end
   str
end

def _details(workbook,worksheet,updates=nil)
     excel=WIN32OLE.new('Excel.Application')  
     
     return err("Workbook #{workbook} does not exist") unless File.exist?(workbook+'.xls')     
     book=excel.Workbooks.Open(DIRECTORY+'/' + workbook)  
     
     total_sheets = book.Worksheets.Count
     return err("Worksheet #{worksheet} does not exist") unless (1..total_sheets).to_a.include?(worksheet.to_i)   
     
     sheet = book.Worksheets(worksheet.to_i)     
     str=''

    # Display the available sheets within the book
     (1..book.Worksheets.Count).each do |i|          
          if i == worksheet.to_i
             str+="| #{book.Worksheets(i).Name}"
          else   
             str+="|<a href=\"/#{workbook}/#{i}\">#{book.Worksheets(i).Name}</a>"
          end          
     end   
     str+=' |<br/>'
     
     # Update cells if any have values suplied by the user
     updates.each_pair{|k,v| sheet.range("#{k}").Value=v} unless updates.nil?

    # Range of Spreadsheet
     rows=nil
     if params[:start].nil? or params[:ending].nil?
        rows = sheet.UsedRange.Value
     else
        rows = sheet.Range("#{params[:start]}:#{params[:ending]}").Value
     end
     
    range=sheet.UsedRange.Address 
    start_col, start_row, end_col, end_row=range.gsub(':','').split('$')[1,4]     
     
     # The Spreadsheet data itself
     unless rows.nil?
       str+="<h3>#{sheet.Name}</h3>"
       str+='<table border="1" cellpadding="2" cellspacing="0">'
       
       # Headers
       str+='<th></th>'
       (start_col..end_col).each{|c|str+="<th>#{c}</th>"}
             
       idx=1             
       (1...start_row.to_i).each{|i|str+="<tr><td>#{idx}</td></tr>"; idx+=1}
       
       rows.each {|r|          
          str+="<tr><td>#{idx}</td>"          
          r.each_with_index{|d,j|str+="<td>#{d}</td>"}
          str+='</tr>'
          idx+=1
       } 
       str+='</table>'
     end
     
    kill_excel_automation_processes
    str
end
#------------------------------------------------------------------------------------------------------------------------
# http://localhost:4567/book/sheet/range/A1/B4
get '/:workbook/:worksheet/range/:start/:ending'  do
   str=_header('Display Range',params)
   str+=_details(params['workbook'],params['worksheet'],nil) 
   str
end
  
# http://localhost:4567/actor/1/setcell?A1=5&B2=%22Hello%20World%22
get '/:workbook/:worksheet/setcell'  do
   str= _header('Set Cell',params)  
  changed_values = params
  b=changed_values.delete('workbook')
  s=changed_values.delete('worksheet')        
  str+=_details(b,s,changed_values) 
  str
end
  
# http://localhost:4567/book/sheet
get '/:workbook/:worksheet'  do
   str= _header('Spreadsheets',params)
   str+=_details(params['workbook'],params['worksheet'],nil) 
   str
 end  
 
 # http://localhost:4567
 get '/'  do
   str= _header('Spreadsheets',nil)
   Dir.entries(DIRECTORY).each do |f|
      if f=~/.*\.xls/
        str+="<a href=\"#{f.gsub('.xls','')}/1\">"
        str+="#{f}</a><br/>" 
      end  
   end   
   str
end