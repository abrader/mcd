#!/usr/bin/env ruby

require 'csv'
require 'rubygems'
require 'roo'

FILENAME = "/Users/abrader/Downloads/silam_symbol_masterlist.csv"
EXCEL_FILE = "SILAM_Symbol_MasterList.xlsx"

class Silam
  
  attr_accessor :quant_method, :common_name, :gi_number, :approved_symbol, :hgnc_link
  
  def initialize(qm, cm, gi, as, hl)
    @quant_method = qm
    @common_name = cm
    @gi_number = gi
    @approved_symbol = as
    @hgnc_link = hl
  end
  
  # def self.get_list
  #   silam_array = Array.new
  #   count = 0
  #   CSV.foreach(FILENAME) do |row|
  #     if count == 0
  #       #Do Nothing since first row is headers
  #     else
  #       silam_array << Silam.new(row[0], row[1], row[2], row[3])
  #     end
  #     count += 1
  #   end
  #   silam_array
  # end
  # 
  # def self.roo_list
  #   silam_array = Array.new
  #   count = 0
  #   oo = Excelx.new(EXCEL_FILE)
  #   oo.default_sheet = oo.sheets.first
  #   oo.formula(1,'D')
  #   #puts oo.to_yaml
  # end
  
  def self.goog_list
    silam_array = Array.new
    oo = Google.new("0Ao_C0TS02CV1dDYtSWYwM2k3US1POG84S1hrZUpEM2c")
    oo.default_sheet = oo.sheets.first
    count = oo.last_row
    1.upto(count) do |r|
      if oo.formula(r,'D').nil?
        silam_array << Silam.new(oo.cell(r,'A'), oo.cell(r,'B'), oo.cell(r,'C'), oo.cell(r,'D'), nil)
      else
        silam_array << Silam.new(oo.cell(r,'A'), oo.cell(r,'B'), oo.cell(r,'C'), oo.cell(r,'D'), oo.formula(r,'D').split('"')[1])
      end
    end
    silam_array
  end
    
  
end
  

sa = Silam.goog_list

sa.each do |s|
  puts "#{s.approved_symbol} = #{s.hgnc_link}"
end