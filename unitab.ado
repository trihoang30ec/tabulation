* unitab a1 a3, dir(unitab.xlsx) temp(temp.xls)
program unitab
version 15.1
syntax [varlist(default = none)] [, dir(string) temp(string)]
	foreach v of varlist `varlist' {
	di `v'
	tabout `v' using "`temp'", cells(freq col) f(0c 1c) replace
	putexcel set "`dir'", sheet("`v'") modify
	putexcel A1:G50 = "", font(ariel, 12, blue) // set up format cho bảng
	preserve
		insheet using "`temp'", clear double
		export excel using "`dir'", sheet("`v'") firstrow(varlabels) keepcellfmt sheetmodify 
	restore
	}
end
