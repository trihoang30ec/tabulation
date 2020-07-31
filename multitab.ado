* multitab c6a3_1-c6a3_7, sheet(c6) dir(D:\Depocen\Depocen project\2019\plan\indicator\c6)
program multitab
version 15.1
syntax [varlist(default = none)] [, sheet(string) dir(string)]
local max = 0
* Get options
foreach i of varlist `varlist' {
levelsof `i', local(levels)
local dem1 : word count `levels'
if `dem1' > `max' {
local max = `dem1'
di "Số lượng phương án max = `dem1'"
* get label value
local nhan : value label `i'
local v = 0
foreach j of local levels {
local ++v
local option_`v' : label `nhan' `j'
}
local j1 = `v'
forvalues j2 = `j1'/10 {
local ++v
local option_`v' = ""
}
}
}

* get variable label
foreach i of varlist `varlist' {
local lb`i' : variable label `i'
local lb`i' = subinstr("`lb`i''","(","",.)
local lb`i' = subinstr("`lb`i''",")","",.)
local lb`i' = subinstr("`lb`i''","...","",.)
}

* tạo bảng tính
tabm `varlist', matcell(frequency) // bảng tần suất
mata: st_matrix("percentage", (st_matrix("frequency") :/rowsum(st_matrix("frequency")))) // bảng tỷ lệ
putexcel set "`dir'", sheet("`sheet'") modify
* get label for frequency table
putexcel A1:G50 = "", font(ariel, 12, blue) // set up format cho bảng
local col_name  B C D E F G H I J K // maximum 10 label values
local j = 0
foreach col1 of local col_name {
local ++j
putexcel `col1'1 = ("`option_`j''") // tên cột
}
local k = 1
local m = `k' + 1
putexcel A`k' = ("Frequency") // tên dòng

foreach i of varlist `varlist' {
local ++k
putexcel A`k' = ("`lb`i''") // label biến
}
putexcel B`m' = matrix(frequency), nformat(number) // làm tròn số

* get labels for percentage table
local k = `k' + 2
local m = `k' + 1
putexcel A`k' = ("Percentage") // tên dòng

local j = 0
foreach col2 of local col_name {
local ++j
putexcel `col2'`k' = ("`option_`j''") // tên cột

}
foreach i of varlist `varlist' {
local ++k
putexcel A`k' = ("`lb`i''") // label biến
}
putexcel B`m' = matrix(percentage), nformat(percent_d2) // làm tròn 2 chữ số
end
