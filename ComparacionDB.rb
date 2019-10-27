if __FILE__ == $0

require 'roo'
require 'xls'
require 'i18n'

#SACAMOS LA INFORMACION DEL DANE
I18n.config.available_locales = :en
workbookDane =Roo::Spreadsheet.open'/home/daniela/Escritorio/veredas.xlsx'
infoDane=Array.new
num_rows_dane = 3
worksheetsDane = workbookDane.sheets
puts "Found #{worksheetsDane.count} worksheets"
worksheetsDane.each do |worksheet1|
  puts "Reading: #{worksheet1}"
workbookDane.sheet(worksheet1).each_row_streaming do |row|

if(workbookDane.cell('B',num_rows_dane))
	departamentoDane = I18n.transliterate(workbookDane.cell('B',num_rows_dane)).upcase
#puts departamentoDane
    end 
 if(workbookDane.cell('C',num_rows_dane))
municipioDane = I18n.transliterate(workbookDane.cell('C',num_rows_dane)).upcase
#puts municipioDane
    end 
 if(workbookDane.cell('E',num_rows_dane))
veredaDane = I18n.transliterate(workbookDane.cell('E',num_rows_dane)).upcase
#puts veredaDane
    end 

   hashe={"DEPARTAMENTO"=> departamentoDane,"MUNICIPIO"=> municipioDane, "VEREDA"=>veredaDane}
   infoDane.push(hashe)
 
    num_rows_dane += 1
  end
  puts "Read #{num_rows_dane} rows" 

end
  puts "Read #{num_rows_dane} rows" 
workbookDane.close
#infoDane.each do |infoDanes|
#infoDanes.each{|key, value| puts "#{key} is #{value}" }

#end


#SACAMOS LA INFORMACION DEL EXCEL A REVISAR
lugarNoticia=Array.new
workbook = Roo::Spreadsheet.open '/home/daniela/Escritorio/Para-ayudar-analizar-Daniela-Sebastian-12-Sep-2019.xlsx'
 num_rows = 2
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"
worksheets.each do |worksheet|
  puts "Reading: #{worksheet}"
 
  workbook.sheet(worksheet).each_row_streaming do |row|
    if(workbook.cell('B',num_rows))
	departamento1 = I18n.transliterate(workbook.cell('B',num_rows)).upcase
#puts departamento1
    end 
 if(workbook.cell('C',num_rows))
municipio1 = I18n.transliterate(workbook.cell('C',num_rows)).upcase
#puts municipio1
    end 
 if(workbook.cell('D',num_rows))
vereda1 = I18n.transliterate(workbook.cell('D',num_rows)).upcase
#puts vereda1
    end 

   hash={"DEPARTAMENTO"=> departamento1,"MUNICIPIO"=> municipio1, "VEREDA"=>vereda1}
   lugarNoticia.push(hash)


#lugarNoticia= { "DEPARTAMENTO" => departamento1,"MUNICIPIO" => municipio1}
depEncontre=0
#empieza a buscar en las veredas que estan}

 
    num_rows += 1
  end
  puts "Read #{num_rows} rows" 
#lugarNoticia.each {|key, value| puts "#{key} is #{value}" }
end
  puts "Read #{num_rows} rows" 

lugarNoticia.each do |lugarNoticias|
lugarNoticias.each{|key, value| puts "#{key} is #{value}" }

end
workbook.close
CSV.open("/home/daniela/Escritorio/result.xlsx", "wb") do |csv|
csv<<["DEPARTAMENTO","EXISTEDEP","MUNICIPIO","EXISTEMUN","VEREDA","EXISTEVER"]


#VAMOS A BUSCAR
i=2

encontreDep=0
encontreMun=0
encontreVere=0
lugarNoticia.each do |lugarNoticias|
infoDane.each do |infoDanes|
if infoDanes["DEPARTAMENTO"]==lugarNoticias["DEPARTAMENTO"] && encontreDep==0
	encontreDep=1
	
   if infoDanes["MUNICIPIO"]==lugarNoticias["MUNICIPIO"] && encontreMun==0
	encontreMun=1
	if infoDanes["VEREDA"]==lugarNoticias["VEREDA"] && encontreVere==0
encontreVere==1
	csv << [lugarNoticias["DEPARTAMENTO"],1,lugarNoticias["MUNICIPIO"],1,lugarNoticias["VEREDA"],1]
 end
puts "hi2"
	
    end
  

end

end
if encontreDep==1 && encontreMun==1 && encontreVere==0
csv << [lugarNoticias["DEPARTAMENTO"],1,lugarNoticias["MUNICIPIO"],1,lugarNoticias["VEREDA"],0]
elsif encontreDep==1 && encontreMun==0 && encontreVere==0
csv << [lugarNoticias["DEPARTAMENTO"],1,lugarNoticias["MUNICIPIO"],0,lugarNoticias["VEREDA"],0]
elsif encontreDep==0 && encontreMun==0 && encontreVere==0
csv << [lugarNoticias["DEPARTAMENTO"],0,lugarNoticias["MUNICIPIO"],0,lugarNoticias["VEREDA"],0]
end

i += 1
encontreDep=0
encontreMun=0
encontreVere=0
end
end
CVS.close
end
