require './domaci.rb'

# 1. Biblioteci se prosledjuje path do excel fajla
t = get_table_from_excel_file('test.xlsx') 
t2 = get_table_from_excel_file('test2.xls') 

# 2. Biblioteka vraca 2d niz sa vrednostima iz tabele
# print t.table 

# 3. Moguće je pristupati redu preko t.row(1), i pristup njegovim elementima po sintaksi niza. 
# print t.row(0)[1] 

# 4. Mora biti implementiran enumerable modul, gde se vraćaju sve ćelije, sa leva na desno. 
# t.eachCell do |cell|
#     puts cell
# end

# 6. [] sintaksa mora da bude obogaćena tako da je moguće pristupati određenim vrednostima. a) i b)
# print t["Kolona2"]
# print t["Kolona2"][1]

# 7. Biblioteka omogućava direktni pristup kolonama, preko istoimenih metoda.
# print t.Kolona3
# print t.Kolona3.sum
# print t.Index
# print t.Index.rn4419

# 8. U fajlu test.xlsx u tabeli se nalazi kljucna rec total, tu se vidi da biliboteka ignorise taj red

# 9. Moguce je sabiranje dve tabele, sve dok su im headeri isti.
# t3 =  t+t2
# print t3.table
# print "\n"

# 10. Moguce je oduzimanje dve tabele, sve dok su im headeri isti.
# t4 = t3 - t2
# print t4.table
# print "\n"
# print t.table

# 11. U fajlu test.xlsx u tabeli prvi red je prazan, tu se vidi da biliboteka ignorise taj red

