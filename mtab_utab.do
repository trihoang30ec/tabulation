program drop _all
* syntax: mtab [varlist], by() sheet() dir() type() format()
/* 
by: cross-tabulation
sheet: sheet name
dir: output file(.xlsx)
type: "percentage" or "frequency" (equivalent to relative frequency table or frequency table, respectively)
format: numeric format such as "number", "percent_d2", etc.
*/
program mtab
version 15
syntax [varlist(default = none)] [, sheet(string) by(string) dir(string) type(string) format(string)]
local byvalues : value label `by'
levelsof `by', local(bylevels)
local bycount : word count `bylevels'
local ++bycount
putexcel set "`dir'", sheet(`sheet') modify
local startcol = 65 /* column A */
foreach i in `bylevels' {
tabm `varlist' if `by' == `i', matcell(frequency) /* ssc install tab_chi */
local ++startcol
mata: st_matrix("percentage", (st_matrix("frequency"):/rowsum(st_matrix("frequency")))) /* create relative frequency matrix "percentage" */
local colname = char(`startcol') /* give a name to a column in Excel */
putexcel `colname'3 = matrix(`type'), nformat(`format') /* put "percentage" or "frequency" matrix */
local colcontent : label `byvalues' `i' 
putexcel `colname'1 = "`colcontent'", left /* add label to Excel columns */
local startcol = `startcol' + `r(c)' - 1
local colname2 = char(`startcol')
putexcel `colname'1:`colname2'1, merge font(ariel, 11, black) txtwrap hcenter
}

local ++startcol
tabm `varlist', matcell(frequency) /* "General" column */
mata: st_matrix("percentage", (st_matrix("frequency"):/rowsum(st_matrix("frequency"))))
local colname = char(`startcol')
putexcel `colname'3 = matrix(`type'), nformat(`format')
putexcel `colname'1 = "Chung", left
local startcol = `startcol' + `r(c)' - 1
local colname2 = char(`startcol')
putexcel `colname'1:`colname2'1, merge font(ariel, 11, black) txtwrap hcenter
local colname = char(`startcol')
local row = 2
foreach rowname in `varlist' {
local ++row
local rowcontent : variable label `rowname'
putexcel A`row' = "`rowcontent'", left
}
foreach varlevels in `varlist' {
labellist `varlevels'
}
local d = 0
local labelvalues 
foreach varlabelvalues in `r(values)' {
local ++d
local labelvalues`d' : word `d' of `r(values)'
local comma ","
local labelvalues `labelvalues' `labelvalues`d'' `comma'
di "`labelvalues'"
}
local labelvalues = strreverse(subinstr(strreverse("`labelvalues'"), ",", "", 1))
di "`labelvalues'"
di "`varlevels'"
mata: labelmatrix = st_vlmap(st_varvaluelabel("`r(varlist)'"), (`labelvalues')) /* get value labels */
mata: valuecolname = J(1, `bycount', labelmatrix) /* create a (1 x n) row vector of the short sequence repeated "bycount" times */
mata: valuecolname
preserve
clear
set obs 1
getmata (var*) = valuecolname /* get string matrix from mata */
export excel using "`dir'", sheet(`sheet') sheetmodify keepcellfmt cell(B2)
restore
putexcel A1:`colname'`row', font(ariel, 11, black)
putexcel A1:`colname'2, border(bottom, medium)
putexcel A`row':`colname'`row', border(bottom, medium)
end

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

* Examples:
utab a1 a4 b4 b8 b9 b15 d5 d11 d13, by(tinh) dir(percentage_table.xlsx) temp(temp.xls) type(percentage) format(percent)
mtab b1_01_b-b1_10_b, by(tinh) sheet(b1) dir(percentage_table.xlsx) type(percentage) format(percent_d2)
mtab b2_1-b2_8, by(tinh) sheet(b2) dir(percentage_table.xlsx) type(percentage) format(percent_d2)








