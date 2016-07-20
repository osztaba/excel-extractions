require 'roo-xls'
xls = Roo::Spreadsheet.open('/Users/oliviersztaba/Documents/LgAg_030401-040229.xls')
 
 index = 0
 rows = (xls.sheet(0).column(5).count {|cell| cell != nil}) + 1 
 rate_schedule = nil
 season = nil
 time_of_use_period = nil
 demand_charge = nil
 energy_charge = nil
 CSV.open("LgAg_030401-040229test.csv", "w") do |csv|

 while index < rows do
   unless xls.sheet(0).row(index)[0] == nil 
     rate_schedule = xls.sheet(0).row(index)[0]
   end
   
   unless xls.sheet(0).row(index)[3] == nil 
     season = xls.sheet(0).row(index)[3]
   end
     
    unless xls.sheet(0).row(index)[4] == nil  
      time_of_use_period = xls.sheet(0).row(index)[4]
      if time_of_use_period == "-"
        time_of_use_period.gsub!('-', '0')
    end
  end
    
    unless xls.sheet(0).row(index)[5] == nil  
       demand_charge =  xls.sheet(0).row(index)[5]
       if demand_charge == "-"
         demand_charge.gsub!('-', '0')
    end
  end
    
    unless xls.sheet(0).row(index)[6] == nil  
       energy_charge =  xls.sheet(0).row(index)[6]
       if energy_charge == "-"
         energy_charge.gsub!('-', '0')
    end
  end
    
   csv << [rate_schedule, season, time_of_use_period, demand_charge, energy_charge]
 
 index += 1
 end
end


 
 
 
 
 
 
 