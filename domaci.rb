require 'roo'
require 'spreadsheet'

# Za dinamicko kreiranje metoda
def add_method(c, m, &b)
  c.class_eval {
    define_method(m.to_s, &b)
  }
end

# Medjukorak-kreiranje klase Table, koja ce sadrzati dobijenu matricu u polju table; dodavanje svih potrebnih funkcija
class Table
  include Enumerable
  attr_accessor :table, :transp_table

  def initialize(matrix)
    @table = matrix # originalna tabela, po redovima
    @transp_table = transpose # transponovana tabela, po kolonama (prva kolona iz originalne tabele je sad prvi red itd.)
    direct_approach_to_columns # omogucava korisniku opciju: t.kolona1, t.kolona2...direktan pristup koloni po njenom nazivu
    allow_getting_row_by_key # omogucava korisniku opciju: t.Index.rn2310 sto ce vratiti red studenta cija je vrednost u koloni Index  =rn2310
  end

  def row(num)
    return @table[num]
  end

  def eachCell(&block)
    @table.each(&block)
  end

  def transpose
    @table.first.zip(*@table[1..-1])
  end

  def [](column_name)
    chash = create_column_hash
    return chash[column_name]
  end

  def +(t2) # t2 je takodje objekat klase Table
    cnt = 0
    @table[0].each do |val|
      if val != t2.table[0][cnt]
        print "Zaglavlja tabela se razlikuju. Tabele ne mogu da se saberu."
        return nil
      end
      cnt += 1
    end
    mat = @table
    mat += t2.table[1..-1]
    t1 = Table.new(mat)
    return t1
  end

  def -(t2) # t2 je takodje objekat klase Table
    cnt = 0
    @table[0].each do |val|
      if val != t2.table[0][cnt]
        print "Zaglavlja tabela se razlikuju. Tabele ne mogu da se saberu."
        return nil
      end
      cnt += 1
    end
    mat = @table
    mat -= t2.table[1..-1]
    t1 = Table.new(mat)
    return t1
  end  

  def to_s
    return @table
  end
  
  private 
  def create_column_hash
    h = Hash.new{ |h,k| h[k]=[] }
    @transp_table.each do |col|
      flag_for_total = 0 
      key = ""
      col.each do |value| 
        if flag_for_total == 0
          key = value
          flag_for_total = 1
        else
          h[key] << value
        end 
      end
    end
    return h
  end

  def direct_approach_to_columns
    cnt = 1
    @table[0].each do |cell|
      cnt2 = cnt-1
      name = @table[0][cnt2]
      transp = @transp_table 
      add_method(Table, name) {
        return transp[cnt2]
      }
      cnt += 1
    end  
  end

  def allow_getting_row_by_key
    #  napraviti metodu za svaki element iz nulte kolone
    cnt = 1
    @transp_table[0].slice(1..-1).each do |val|
      help_table = @table
      cnt2 = cnt
      add_method(Array, val.to_s){
        return help_table[cnt2]
      }
      cnt += 1
    end
  end

end

class Integer
  def include?(str)
    return false
  end
end

class Float
  def include?(str)
    return false
  end
end

class Array
  def sum
    sum = 0
    self.each do |value|
      sum += value.to_i
    end
    return sum
  end
end

#<Otvaranje excel fajla, trazenje tabele i konvertovanje u dvodimenzionalni niz>

def get_table_from_excel_file(excel_file)
  if excel_file.end_with? ".xlsx"
    return get_table_from_xlsx_file(excel_file)
  end
  if excel_file.end_with? ".xls"
    return get_table_from_xls_file(excel_file)
  end 
  print("Greska. Biblioteka podrzava samo .xls i .xlsx fajlove") 
  return nil
end

def get_table_from_xlsx_file(excel_file)
  workbook = Roo::Spreadsheet.open(excel_file, {:expand_merged_ranges => true})
  worksheets = workbook.sheets
  puts "Found #{worksheets.count} worksheets" 

  dim = Array.new
  arr = Array.new

  table_columns = 0

  worksheets.each do |worksheet|
    puts "Reading: #{worksheet}"
    num_rows = 0
    workbook.sheet(worksheet).each_row_streaming do |row|
    flag_for_total = false
      row_cells = row.map {
        |cell| cell.value
      }
      row_cells.each do |cell| # proveravamo da li red negde sadrzi kljucne reci total ili subtotal, ako ima, kuliramo ceo red
        if cell != nil
          if cell.include? "total" or cell.include? "subtotal" 
            flag_for_total = true
          end
        end 
      end
      row_cells.each do |cell|
        if cell != nil and flag_for_total == false
          arr.push(cell)
          table_columns += 1
        end 
      end
      num_rows += 1
      if table_columns > 0
        dim.push(table_columns)
      end  
      table_columns = 0
    end
    puts "Read #{num_rows} rows" 
  end

  col_number = dim[0]
  cnt = 0
  help_array = Array.new
  matrix = Array.new

  arr.each do |element|
    help_array.append(element)
    cnt += 1 
    if cnt == col_number
      cnt = 0
      help_array_clone = help_array.clone
      matrix.push(help_array_clone)
      help_array.clear
    end
  end
  t1 = Table.new(matrix)
  return t1
end

def get_table_from_xls_file(excel_file)
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet.open excel_file
  worksheets = book.worksheets
  puts "Found #{worksheets.count} worksheets" 

  dim = Array.new
  arr = Array.new

  table_columns = 0

  worksheets.each do |worksheet|
    puts "Reading: #{worksheet}"
    num_rows = 0
    worksheet.each do |row|
      flag_for_total = false
      if row != nil 
        row_cells = row.map {
            |cell| cell
        }
      end
      row_cells.each do |cell| # proveravamo da li red negde sadrzi kljucne reci total ili subtotal, ako ima, kuliramo ceo red
        if cell != nil
          if cell.include? "total" or cell.include? "subtotal" 
            flag_for_total = true
          end
        end 
      end
      row_cells.each do |cell|
        if cell != nil and flag_for_total == false
          arr.push(cell)
          table_columns += 1
        end 
      end
      num_rows += 1
      if table_columns > 0
        dim.push(table_columns)
      end  
      table_columns = 0
    end
    puts "Read #{num_rows} rows" 
  end

  col_number = dim[0]
  cnt = 0
  help_array = Array.new
  matrix = Array.new

  arr.each do |element|
    help_array.append(element)
    cnt += 1 
    if cnt == col_number
      cnt = 0
      help_array_clone = help_array.clone
      matrix.push(help_array_clone)
      help_array.clear
    end
  end
  t1 = Table.new(matrix)
  return t1
end

