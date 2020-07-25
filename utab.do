* syntax: utab [varlist], by() dir() temp() type() format()
/* 
by: cross-tabulation
dir: output file(.xlsx)
temp: temporary file (.xls) from tabout
type: "percentage" or "frequency" (equivalent to relative frequency table or frequency table, respectively)
format: numeric format such as "number", "percent_d2", etc.
*/
program utab
version 15.1
syntax [varlist(default = none)] [, by(string) dir(string) temp(string) type(string) format(string)]
	foreach v in `varlist' {
	di `v'
	tabout `v' `by' using "`temp'", cells(col) f(1c) replace
	putexcel set "`dir'", sheet("`v'_`by'") modify
	preserve
		insheet using "`temp'", clear double
		drop in 1
		drop in 2
		export excel using "`dir'", sheet("`v'_`by'") keepcellfmt sheetmodify 
	restore
	levelsof `v', local(bylevelsrow)
	local bycountrow : word count `bylevelsrow'
	local rowname = `bycountrow' + 2
	tab `v' `by', matcell(frequency)
	mata: st_matrix("percentage", (st_matrix("frequency"):/colsum(st_matrix("frequency"))))
	putexcel B2 = matrix(`type'), nformat(`format')
	mata: st_matrix("`type'", colsum(st_matrix("`type'")))
	putexcel B`rowname' = matrix(`type'), nformat(`format')
	tab `v', matcell(frequency)
	mata: st_matrix("percentage", (st_matrix("frequency"):/colsum(st_matrix("frequency"))))
	levelsof `by', local(bylevels)
	local bycount : word count `bylevels'
	local startcol = 65 + `bycount' + 1
	local colname = char(`startcol')
	putexcel `colname'2 = matrix(`type'), nformat(`format')
	mata: st_matrix("`type'", colsum(st_matrix("`type'")))
	putexcel `colname'`rowname' = matrix(`type'), nformat(`format')
	putexcel A1:`colname'`rowname', font(ariel, 11, black) // set up format cho bảng
	putexcel A`rowname':`colname'`rowname', border(bottom, medium)
	putexcel A1:`colname'1, border(bottom, medium)
	}
end

* Example: 
utab a1 a4 b4 b8 b9 b15 d5 d11 d13, by(tinh) dir(frequency_table.xlsx) temp(temp.xls) type(frequency) format(number)
utab a1 a4 b4 b8 b9 b15 d5 d11 d13, by(tinh) dir(percentage_table.xlsx) temp(temp.xls) type(percentage) format(percent_d2)
