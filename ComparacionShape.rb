if __FILE__ == $0

require 'roo'
require 'xls'
require 'i18n'

#SACAMOS LA INFORMACION DEL DANE
I18n.config.available_locales = :en
#SACAMOS LA INFORMACION DEL EXCEL A REVISAR
lugarNoticia1=Array.new
lugarNoticiaRevision=Array.new
lugarNoticiaSoft=Array.new
workbook = Roo::Spreadsheet.open '/home/daniela/Escritorio/Para-ayudar-analizar-Daniela-Sebastian-12-Sep-2019.xlsx'
 num_rows = 2
vereda3_11="nil"
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
		#LUGARES DE LA REVISION MANUAL
		if(workbook.cell('E',num_rows))
			departamento2 = I18n.transliterate(workbook.cell('E',num_rows)).upcase
		#puts departamento1
		    end 
		if(!workbook.cell('E',num_rows))
			departamento2 = departamento1
		#puts departamento1
		    end 
		 if(workbook.cell('F',num_rows))
		municipio2 = I18n.transliterate(workbook.cell('F',num_rows)).upcase
		#puts municipio1
		    end 
		 if(!workbook.cell('F',num_rows))
		municipio2= municipio1
		    end 
		 if(workbook.cell('H',num_rows))
		vereda2 = I18n.transliterate(workbook.cell('H',num_rows)).upcase
		#puts vereda1
		    end 
		 if(!workbook.cell('H',num_rows))
		vereda2 =vereda1
		#puts vereda1
		    end 


		#LUGAR DEL SOFTWARE
		 if(workbook.cell('M',num_rows))
		vereda3_11 = I18n.transliterate(workbook.cell('M',num_rows)).upcase
		#puts vereda1
		    end 
		if(workbook.cell('O',num_rows))
			departamento3_1 = I18n.transliterate(workbook.cell('O',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('P',num_rows))
		municipio3_1 = I18n.transliterate(workbook.cell('P',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('Q',num_rows))
		vereda3_1 = I18n.transliterate(workbook.cell('Q',num_rows)).upcase
		#puts vereda1
		    end 

		if(workbook.cell('S',num_rows))
			departamento3_2 = I18n.transliterate(workbook.cell('S',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('T',num_rows))
		municipio3_2 = I18n.transliterate(workbook.cell('T',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('U',num_rows))
		vereda3_2 = I18n.transliterate(workbook.cell('U',num_rows)).upcase
		#puts vereda1
		    end 

		if(workbook.cell('W',num_rows))
			departamento3_3 = I18n.transliterate(workbook.cell('W',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('X',num_rows))
		municipio3_3 = I18n.transliterate(workbook.cell('X',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('Y',num_rows))
		vereda3_3 = I18n.transliterate(workbook.cell('Y',num_rows)).upcase
		#puts vereda1
		    end 



if(workbook.cell('AA',num_rows))
			departamento3_4 = I18n.transliterate(workbook.cell('AA',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('AB',num_rows))
		municipio3_4 = I18n.transliterate(workbook.cell('AB',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('AC',num_rows))
		vereda3_4 = I18n.transliterate(workbook.cell('AC',num_rows)).upcase
		#puts vereda1
		    end 


if(workbook.cell('AE',num_rows))
			departamento3_5 = I18n.transliterate(workbook.cell('AE',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('AF',num_rows))
		municipio3_5 = I18n.transliterate(workbook.cell('AF',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('AG',num_rows))
		vereda3_5 = I18n.transliterate(workbook.cell('AG',num_rows)).upcase
		#puts vereda1
		    end 


if(workbook.cell('AI',num_rows))
			departamento3_6 = I18n.transliterate(workbook.cell('AI',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('AJ',num_rows))
		municipio3_6 = I18n.transliterate(workbook.cell('AJ',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('AK',num_rows))
		vereda3_6 = I18n.transliterate(workbook.cell('AK',num_rows)).upcase
		#puts vereda1
		    end 

if(workbook.cell('AM',num_rows))
			departamento3_7 = I18n.transliterate(workbook.cell('AM',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('AN',num_rows))
		municipio3_7 = I18n.transliterate(workbook.cell('AN',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('AO',num_rows))
		vereda3_7 = I18n.transliterate(workbook.cell('AO',num_rows)).upcase
		#puts vereda1
		    end 

if(workbook.cell('AQ',num_rows))
			departamento3_8 = I18n.transliterate(workbook.cell('AQ',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('AR',num_rows))
		municipio3_8 = I18n.transliterate(workbook.cell('AR',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('AS',num_rows))
		vereda3_8 = I18n.transliterate(workbook.cell('AS',num_rows)).upcase
		#puts vereda1
		    end 

if(workbook.cell('AU',num_rows))
			departamento3_9 = I18n.transliterate(workbook.cell('AU',num_rows)).upcase
		#puts departamento1
		    end 
		 if(workbook.cell('AV',num_rows))
		municipio3_9 = I18n.transliterate(workbook.cell('AV',num_rows)).upcase
		#puts municipio1
		    end 
		 if(workbook.cell('AW',num_rows))
		vereda3_9 = I18n.transliterate(workbook.cell('AW',num_rows)).upcase
		#puts vereda1
		    end 
                  hash={"DEPARTAMENTO"=> departamento1,"MUNICIPIO"=> municipio1, "VEREDA"=>vereda1}
		  lugarNoticia1.push(hash)
		  hash2={"DEPARTAMENTO"=> departamento2,"MUNICIPIO"=> municipio2, "VEREDA"=>vereda2}
		  lugarNoticiaRevision.push(hash2)
		  hash3={"DEPARTAMENTO"=> departamento3_1,"MUNICIPIO"=> municipio3_1, "VEREDA"=>vereda3_1}
		  lugarNoticiaSoft.push(hash3)
		  hash3={"DEPARTAMENTO"=> departamento3_2,"MUNICIPIO"=> municipio3_2, "VEREDA"=>vereda3_2}
		  lugarNoticiaSoft.push(hash3)
		  hash3={"DEPARTAMENTO"=> departamento3_3,"MUNICIPIO"=> municipio3_3, "VEREDA"=>vereda3_3}
		  lugarNoticiaSoft.push(hash3)
  hash3={"DEPARTAMENTO"=> departamento3_4,"MUNICIPIO"=> municipio3_4, "VEREDA"=>vereda3_4}
		  lugarNoticiaSoft.push(hash3)
  hash3={"DEPARTAMENTO"=> departamento3_5,"MUNICIPIO"=> municipio3_5, "VEREDA"=>vereda3_5}
		  lugarNoticiaSoft.push(hash3)
  hash3={"DEPARTAMENTO"=> departamento3_6,"MUNICIPIO"=> municipio3_6, "VEREDA"=>vereda3_6}
		  lugarNoticiaSoft.push(hash3)
  hash3={"DEPARTAMENTO"=> departamento3_7,"MUNICIPIO"=> municipio3_7, "VEREDA"=>vereda3_7}
		  lugarNoticiaSoft.push(hash3)
  hash3={"DEPARTAMENTO"=> departamento3_8,"MUNICIPIO"=> municipio3_8, "VEREDA"=>vereda3_8}
		  lugarNoticiaSoft.push(hash3)
  hash3={"DEPARTAMENTO"=> departamento3_9,"MUNICIPIO"=> municipio3_9, "VEREDA"=>vereda3_9}
		  lugarNoticiaSoft.push(hash3)

	#lugarNoticia= { "DEPARTAMENTO" => departamento1,"MUNICIPIO" => municipio1}
		  depEncontre=0
	#empieza a buscar en las veredas que estan}

	 
	          num_rows += 1
	  end
	  puts "Read #{num_rows} rows" 
	#lugarNoticia.each {|key, value| puts "#{key} is #{value}" }
end
puts "Read #{num_rows} rows" 

	#lugarNoticia.each do |lugarNoticias|
	#lugarNoticias.each{|key, value| puts "#{key} is #{value}" }

	#end
workbook.close
encontreDep=0
	encontreMun=0
	encontreVere=0
CSV.open("/home/daniela/Escritorio/result.xlsx", "wb") do |csv|
	csv<<["DEPARTAMENTO","ISEQUAL","MUNICIPIO","ISEQUAL","VEREDA","ISEQUAL"]

	#VAMOS A BUSCAR
	i=2

	encontreDep=0
	encontreMun=0
	encontreVere=0
	lugarNoticia1.each do |lugarNoticias|
		if lugarNoticias["VEREDA"]== vereda3_11
		encontreVere=1
		end
		lugarNoticiaSoft.each do |infoSoft|
			if infoSoft["DEPARTAMENTO"]==lugarNoticias["DEPARTAMENTO"]
				encontreDep=1
			end
			if infoSoft["MUNICIPIO"]==lugarNoticias["MUNICIPIO"]
				encontreMun=1
			end
			if infoSoft["VEREDA"]==lugarNoticias["VEREDA"] 
			encontreVere=1
			 end
	
	
	         end

		 csv << [lugarNoticias["DEPARTAMENTO"],encontreDep,lugarNoticias["MUNICIPIO"],encontreMun,lugarNoticias["VEREDA"],encontreVere]
		i += 1
		encontreDep=0
		encontreMun=0
		encontreVere=0
	end


	##
	csv <<["DEP_CAMBIA","ISEQUAL","MUN_CAMBIA","ISEQUAL","VEREDA_CAMBIA","ISEQUAL"]
	lugarNoticiaRevision.each do |lugarNoticias|
		if lugarNoticias["VEREDA"]== vereda3_11
		encontreVere=1
		end
		lugarNoticiaSoft.each do |infoSoft|
			if infoSoft["DEPARTAMENTO"]==lugarNoticias["DEPARTAMENTO"]
				encontreDep=1
			end
			if infoSoft["MUNICIPIO"]==lugarNoticias["MUNICIPIO"]
				encontreMun=1
			end
			if infoSoft["VEREDA"]==lugarNoticias["VEREDA"] 
			encontreVere=1
			 end
	         end
	
	    
          

	 csv << [lugarNoticias["DEPARTAMENTO"],encontreDep,lugarNoticias["MUNICIPIO"],encontreMun,lugarNoticias["VEREDA"],encontreVere]
	i += 1
	encontreDep=0
	encontreMun=0
	encontreVere=0
	end

end
CSV.close
end
